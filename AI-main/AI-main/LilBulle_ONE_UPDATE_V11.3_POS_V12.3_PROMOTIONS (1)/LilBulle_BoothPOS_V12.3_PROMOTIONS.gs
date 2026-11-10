/** lilbulle POS V12.1.15 | 2026-10-30 | User-confirmed LINE summaries only; no automatic sale or close notifications.
 * Replace this extension in the SAME Apps Script project. API schema remains 11.4.1. */
var LB_APP_VERSION='12.3.0';
var LB_LEDGER='LilbullePOS_V11';
function lbDay_(){return Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd');}
// Sheets may automatically parse yyyy-MM-dd cells as native Date values.
function lbNormalizeDay_(value){
 if(value instanceof Date)return Utilities.formatDate(value,'Asia/Bangkok','yyyy-MM-dd');
 return String(value||'').trim();
}
function lbSetup_(){var ss=posSetup_();posGetSheet_(ss,LB_LEDGER,['ID','RequestID','Type','IssueNo','Day','At','Staff','PayloadJSON','Notify','NotifyError']);return ss;}
function lbEvents_(){return posRows_(lbSetup_().getSheetByName(LB_LEDGER)).map(function(r){return {id:String(r[0]),requestId:String(r[1]),type:String(r[2]),issue:String(r[3]),day:lbNormalizeDay_(r[4]),at:r[5] instanceof Date?r[5].toISOString():String(r[5]),staff:String(r[6]),data:JSON.parse(r[7]),notify:String(r[8]),error:String(r[9]||'')};});}
function lbQty_(n,zero){if(n===''||n===null||n===undefined||!Number.isInteger(Number(n))||Number(n)<(zero?0:1)||Number(n)>100000)throw Error('กรอกจำนวนเต็มที่ถูกต้อง');return Number(n);}
function lbKey_(sku,exp){return sku+'|'+normalizeExpKey(exp);}
function lbSku_(full){var found='';Object.keys(ITEMS_MAP).forEach(function(k){if(normalizeData(ITEMS_MAP[k].full)===normalizeData(full))found=ITEMS_MAP[k].kw;});if(!found)throw Error('สินค้าในใบเบิกยังไม่จับคู่ SKU: '+full);return found;}
function lbCatalog_(){return posInitV1040().products.map(function(p){p.imageSource=String(p.imageSource||p.image||'');p.image='';p.imageVersion=p.imageSource||('original-v1-'+p.sku);return p;});}
function lbState_(issue,events){
 var ss=lbSetup_(),rows=posRows_(ss.getSheetByName('EventIssue')),selected=rows.filter(function(r){return String(r[1])===issue;});
 if(!selected.length)throw Error('ไม่พบใบเบิกบูธ');
 var es=(events||lbEvents_()).filter(function(e){return e.issue===issue;}),lots={},voids={};
 function makeLot(sku,exp){var key=lbKey_(sku,exp);if(!lots[key])lots[key]={key:key,sku:sku,exp:formatExpForSheet_(exp),issued:0,returned:0,sold:0,out:0,pending:0,remaining:0};return lots[key];}
 var receives=es.filter(function(e){return e.type==='RECEIVE';});
 // After the initial physical booth count, freeze POS opening balance. Subsequent EventIssue changes
 // are ONLY warehouse records and must not be counted again in POS when staff registers restock.
 if(receives.length && Array.isArray(receives[0].data.counts)){
   receives[0].data.counts.forEach(function(x){var parts=String(x.key||'').split('|'),sku=parts[0],exp=parts.slice(1).join('|');if(!sku||!exp)throw Error('ข้อมูลล็อตตรวจรับเดิมไม่ครบ');var original=selected.filter(function(r){return lbKey_(lbSku_(r[3]),r[4])===x.key;})[0];makeLot(sku,original?original[4]:exp).issued+=Number(x.actual)||0;});
 }else{
   selected.forEach(function(r){makeLot(lbSku_(r[3]),r[4]).issued+=Number(r[5])||0;});
 }
 var receivedRestocks={};es.filter(function(e){return e.type==='RESTOCK_RECEIVE';}).forEach(function(e){receivedRestocks[e.data.transferId]=true;});
 var transfers=es.filter(function(e){return e.type==='RESTOCK';});
 transfers.forEach(function(e){(e.data.lines||[]).forEach(function(l){var lot=makeLot(l.sku,l.exp),qty=Number(l.qty)||0;if(receivedRestocks[e.id])lot.issued+=qty;else lot.pending+=qty;});});
 // Warehouse physical return is shown only for audit; it never reads or changes POS sales totals.
 var ret=ss.getSheetByName('EventReturn');if(ret)posRows_(ret).forEach(function(r){if(String(r[2])!==issue)return;var key=lbKey_(lbSku_(r[4]),r[5]);if(lots[key])lots[key].returned+=Number(r[6])||0;});
 es.filter(function(e){return e.type==='VOID';}).forEach(function(e){voids[e.data.billId]=true;});
 var sales=es.filter(function(e){return e.type==='SALE'&&!voids[e.id];});
 sales.forEach(function(e){(e.data.lines||[]).forEach(function(l){var lot=makeLot(l.sku,l.exp);lot.sold+=Number(l.qty)||0;});});
 es.filter(function(e){return e.type==='OUT';}).forEach(function(e){(e.data.lines||[]).forEach(function(l){makeLot(l.sku,l.exp).out+=Number(l.qty)||0;});});
 Object.keys(lots).forEach(function(k){var l=lots[k];l.remaining=l.issued-l.returned-l.sold-l.out;});
 return {id:issue,name:String(selected[0][2]),location:String(selected[0][8]||''),closed:selected.some(function(r){return String(r[7])==='ปิดแล้ว';}),received:receives.length>0,pendingRestocks:transfers.filter(function(e){return !receivedRestocks[e.id];}),movements:es.filter(function(e){return e.type==='OUT'||e.type==='RESTOCK'||e.type==='RESTOCK_RECEIVE';}),lots:Object.keys(lots).map(function(k){return lots[k];}),sales:sales,events:es,day:lbDay_(),dayClosed:es.some(function(e){return e.type==='CLOSE'&&e.day===lbDay_();})};
}
function lbSoldLot_(issue,sku,exp){var events=lbEvents_().filter(function(e){return e.issue===issue;}),voids={},qty=0;events.filter(function(e){return e.type==='VOID';}).forEach(function(e){voids[e.data.billId]=true;});events.filter(function(e){return e.type==='SALE'&&!voids[e.id];}).forEach(function(e){e.data.lines.forEach(function(l){if(l.sku===sku&&normalizeExpKey(l.exp)===normalizeExpKey(exp))qty+=l.qty;});});return qty;}
function lbSummary_(state,day){var list=state.sales.filter(function(e){return e.day===day;}),money={CASH:0,TRANSFER:0,CARD:0,OTHER:0},products={};list.forEach(function(e){money[e.data.payment]+=e.data.total;e.data.lines.forEach(function(l){if(!products[l.sku])products[l.sku]={sku:l.sku,name:l.name,qty:0,total:0};products[l.sku].qty+=l.qty;products[l.sku].total+=l.qty*l.price;});});Object.keys(money).forEach(function(k){money[k]=Math.round(money[k]*100)/100;});return {day:day,bills:list.length,money:money,total:Math.round(Object.keys(money).reduce(function(a,k){return a+money[k];},0)*100)/100,products:Object.keys(products).map(function(k){return products[k];}),movements:lbMovementSummary_(state,day),sales:list};}
function lbSnapshot_(issue){var events=lbEvents_(),issues={},ss=lbSetup_(),sh=ss.getSheetByName('EventIssue');if(sh)posRows_(sh).forEach(function(r){if(String(r[7])!=='ปิดแล้ว')issues[String(r[1])]={id:String(r[1]),name:String(r[2]),location:String(r[8]||'')};});var state=issue?lbState_(issue,events):null;if(state){state.lineSummaries=state.events.filter(function(e){return e.type==='SUMMARY';}).map(function(e){return {id:e.id,day:e.day,notify:e.notify,notifyError:e.error};});state.closes=state.events.filter(function(e){return e.type==='CLOSE';}).map(function(e){delete e.data.notificationText;return {id:e.id,day:e.day,staff:e.staff,data:e.data,notify:e.notify,notifyError:e.error};});delete state.events;state.sales.forEach(function(e){delete e.data.notificationText;});}return {version:'11.4.1',appVersion:LB_APP_VERSION,notificationPolicy:'MANUAL_SUMMARY',bundlePrices:lbBundlePrices_(),promotions:lbPromotions_(),warehouse:lbWarehouse_(),products:lbCatalog_(),booths:Object.keys(issues).map(function(k){return issues[k];}).reverse(),state:state,summary:state?lbSummary_(state,lbDay_()):null,pendingNotifications:events.filter(function(e){return lbShouldNotify_(e)&&e.notify.indexOf('PENDING')===0&&(!issue||e.issue===issue);}).length};}
function lbAppend_(type,p,data){if(['SALE','VOID'].indexOf(type)>=0)data.notificationPolicy='MANUAL_SUMMARY';var json=JSON.stringify(data);if(json.length>44000)throw Error('ข้อมูลรายการยาวเกินกำหนด');var sh=lbSetup_().getSheetByName(LB_LEDGER),id=type+'-'+Utilities.getUuid();if(type==='SUMMARY'){var next=sh.getLastRow()+1;if(next>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),1);sh.getRange(next,3).clearDataValidations();}sh.appendRow([id,p.requestId,type,p.issueNo||'',lbDay_(),new Date(),(/^=/.test(p.staff)?"'"+p.staff:p.staff),json,type==='SUMMARY'?'PENDING':'NONE','']);SpreadsheetApp.flush();return {success:true,id:id,data:data};}
// Owner-approved policy: both authenticated POS roles have full operational access.
function lbRequireBoss_(role){if(['boss','team'].indexOf(role)<0)throw Error('กรุณาเข้าสู่ระบบ');}
function lbMutate_(action,p,role){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  p.requestId=posText_(p.requestId,120);p.staff=posText_(p.staff,80);p.issueNo=posText_(p.issueNo,100);
  if(!/^[A-Za-z0-9_-]{1,120}$/.test(p.requestId)||!p.staff)throw Error('กรอกพนักงานและรหัสคำขอ');
  var all=lbEvents_(),old=all.filter(function(e){return e.requestId===p.requestId;})[0];
  if(old){if(old.issue!==String(p.issueNo||'')||old.type!==({receive:'RECEIVE',sale:'SALE',void:'VOID',closeDay:'CLOSE'})[action])throw Error('รหัสคำขอซ้ำกับรายการอื่น');return {success:true,id:old.id,duplicate:true,data:old.data,notificationStatus:lbShouldNotify_(old)?old.notify:'MANUAL',notificationPending:lbShouldNotify_(old)&&old.notify.indexOf('PENDING')===0,notificationError:lbShouldNotify_(old)?old.error:''};}
  var s=lbState_(p.issueNo,all);if(s.closed)throw Error('ใบเบิกนี้ปิดงานแล้ว');
  var result;
  if(action==='receive'){
   if(s.received)throw Error('ตรวจรับบูธนี้แล้ว');
   // Existing POS sales cannot be inferred or silently migrated.
   var legacySS=lbSetup_(),legacyBooths=posRows_(legacySS.getSheetByName(POS_SHEETS_.booths)).filter(function(b){return String(b[7])===p.issueNo;}).map(function(b){return String(b[0]);});
   if(posRows_(legacySS.getSheetByName(POS_SHEETS_.bills)).some(function(b){return legacyBooths.indexOf(String(b[2]))>=0&&String(b[8])==='PAID';}))throw Error('ใบเบิกนี้มีบิล POS รุ่นเก่า ต้องตรวจย้ายยอดก่อนใช้งาน');
   if(!Array.isArray(p.counts)||p.counts.length!==s.lots.length)throw Error('ตรวจนับให้ครบทุกล็อต');
   var seen={};p.counts.forEach(function(x){var l=s.lots.filter(function(l){return l.key===x.key;})[0];if(!l||seen[x.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[x.key]=true;if(lbQty_(x.actual,true)!==l.remaining){if(posText_(p.note,500).length<5)throw Error('ยอดตรวจรับต่างจากใบเบิก โปรดระบุเหตุผลอย่างน้อย 5 ตัวอักษร');}});
   result=lbAppend_('RECEIVE',p,{counts:p.counts,note:posText_(p.note,500)});
  }else{
   if(!s.received)throw Error('ตรวจรับสินค้าก่อนเปิดขาย');
   if(s.dayClosed)throw Error('ปิดยอดวันนี้แล้ว เริ่มขายได้ในวันถัดไป');
   if(action==='sale'){
    if(['CASH','TRANSFER','CARD','OTHER'].indexOf(p.payment)<0)throw Error('ช่องทางรับเงินไม่ถูกต้อง');
    if(!Array.isArray(p.items)||!p.items.length||p.items.length>30)throw Error('เลือกสินค้า 1–30 รายการ');
    var catalog=lbCatalog_(),lines=[],total=0;
    p.items.forEach(function(x){var product=catalog.filter(function(c){return c.sku===x.sku;})[0];if(!product)throw Error('ไม่พบสินค้า');var qty=lbQty_(x.qty,false),price=posMoney_(x.price);if(price<=0)throw Error('ราคาต้องมากกว่า 0');var remaining=qty;
     s.lots.filter(function(l){return l.sku===x.sku;}).sort(function(a,b){return parseExpForSort(a.exp)-parseExpForSort(b.exp);}).forEach(function(l){var cut=Math.min(remaining,Math.max(0,l.remaining));if(cut){lines.push({key:l.key,sku:l.sku,exp:l.exp,name:product.name,qty:cut,price:price});l.remaining-=cut;remaining-=cut;}});
     if(remaining){lines.push({key:lbKey_(x.sku,'-'),sku:x.sku,exp:'-',name:product.name,qty:remaining,price:price,unallocated:true});}total+=Math.round(qty*price*100)/100;
    });
    if(p.photoId&&!all.some(function(e){return e.type==='PHOTO'&&e.id===p.photoId&&e.issue===p.issueNo&&e.data.kind==='receipt';}))throw Error('ไม่พบไฟล์แนบ');
    result=lbAppend_('SALE',p,{lines:lines,total:Math.round(total*100)/100,payment:p.payment,note:posText_(p.note,500),photoId:p.photoId||''});
   }else if(action==='void'){
    lbRequireBoss_(role);var bill=s.sales.filter(function(e){return e.id===p.billId;})[0];if(!bill)throw Error('ไม่พบบิลที่ยกเลิกได้');if(bill.day!==lbDay_())throw Error('ยกเลิกได้เฉพาะบิลวันนี้ที่ยังไม่ปิดยอด');if(posText_(p.note,500).length<5)throw Error('ระบุเหตุผลอย่างน้อย 5 ตัวอักษร');result=lbAppend_('VOID',p,{billId:p.billId,note:posText_(p.note,500)});
   }else if(action==='closeDay'){
    var summary=lbSummary_(s,lbDay_()),money={},seenLots={},counts=p.counts;if(!Array.isArray(counts)||counts.length!==s.lots.length)throw Error('นับสินค้าคงเหลือครบทุกล็อต');
    var countResults=counts.map(function(x){var l=s.lots.filter(function(l){return l.key===x.key;})[0];if(!l||seenLots[x.key])throw Error('รายการนับซ้ำหรือไม่ถูกต้อง');seenLots[x.key]=true;return {key:l.key,sku:l.sku,exp:l.exp,expected:l.remaining,actual:lbQty_(x.actual,true)};});
    Object.keys(summary.money).forEach(function(k){if(!p.money||p.money[k]===''||p.money[k]===undefined)throw Error('กรอกเงินรับจริงทุกช่องทาง');money[k]=posMoney_(p.money[k]);});
    var mismatch=countResults.some(function(x){return x.expected!==x.actual;})||Object.keys(money).some(function(k){return money[k]!==summary.money[k];});
    if(mismatch){lbRequireBoss_(role);if(posText_(p.note,500).length<5)throw Error('ยอดต่างกัน กรุณาระบุเหตุผลก่อนปิดยอด');}
    delete summary.sales;result=lbAppend_('CLOSE',p,{summary:summary,counts:countResults,actualMoney:money,mismatch:mismatch,note:posText_(p.note,500)});
   }else throw Error('คำสั่งไม่รองรับ');
  }
  // Sales committed in one append; delivery failure must never turn into another sale.
  result.notificationStatus='MANUAL';result.notificationPending=false;
  return result;
 }finally{lock.releaseLock();}
}
// V12.1.15: New marked bill notifications and daily close summaries.
// V12.1.15: old skipped or unmarked sale notifications are never reactivated.
function lbShouldNotify_(e){return !!e&&e.type==='SUMMARY'&&e.data&&e.data.notificationPolicy==='MANUAL_SUMMARY';}
function lbRowShouldNotify_(row){try{return lbShouldNotify_({type:String(row[2]),data:JSON.parse(String(row[7]))});}catch(ignore){return false;}}
function lbNotifyText_(event){
 if(event.type==='SUMMARY'&&lbShouldNotify_(event))return String(event.data.manualText||'');
 var state=lbState_(event.issue),head='บูธ '+String(state.name).slice(0,100)+' วันที่ '+event.day+'\n';
 if(event.type==='CLOSE')return lbDailyCloseText_(event,state.name);
 if(!lbShouldNotify_(event))return '';
 if(event.type==='VOID')return head+'ยกเลิกบิล '+event.data.billId+'\nผู้ทำรายการ '+event.staff+'\nเหตุผล '+String(event.data.note||'').slice(0,500)+'\nเลขรายการ '+event.id;
 var d=event.data,payNames={CASH:'เงินสด',TRANSFER:'โอนเงิน',CARD:'บัตร',OTHER:'อื่น ๆ'};
 return head+'บันทึกขายสำเร็จ\n'+d.lines.map(function(l,i){return (i+1)+'. '+l.name+' ×'+l.qty+' @'+Number(l.price).toFixed(2)+' = '+(l.qty*l.price).toFixed(2)+' บาท';}).join('\n')+'\nยอดบิล '+Number(d.total).toFixed(2)+' บาท · '+(payNames[d.payment]||d.payment)+'\nผู้ขาย '+event.staff+(d.note?'\nหมายเหตุ '+String(d.note).slice(0,500):'')+'\nยอดขายวันนี้ '+lbSummary_(state,event.day).total.toFixed(2)+' บาท\nเลขบิล '+event.id;
}
function lbDailyCloseText_(event,name){
 var d=event.data,sum=d.summary,names={balm:'Baby Rich Balm',lotion:'Organic Lotion',foam:'Baby Foam',mozzie:'Mozzie Guard',spray:'Sanitizer Spray'},balance={},payNames={CASH:'เงินสด',TRANSFER:'โอนเงิน',CARD:'บัตร',OTHER:'อื่น ๆ'};
 var money=function(n){return Number(n||0).toFixed(2);};
 var lines=['บูธ '+String(name).slice(0,100),'วันที่ '+event.day+' · สรุปปิดยอดรายวัน','ยอดขาย '+money(sum.total)+' บาท / '+sum.bills+' บิล'];
 (sum.products||[]).filter(function(p){return Number(p.qty)!==0||Number(p.total)!==0;}).forEach(function(p){names[p.sku]=String(p.name).slice(0,60);lines.push('• '+names[p.sku]+' ขาย '+p.qty+' ชิ้น = '+money(p.total)+' บาท');});
 Object.keys(payNames).filter(function(k){return Number(d.actualMoney[k]||0)!==0||Number(sum.money[k]||0)!==0;}).forEach(function(k){lines.push(payNames[k]+': รับจริง '+money(d.actualMoney[k])+' / ตามบิล '+money(sum.money[k]));});
 var moves=(sum.movements||[]).filter(function(m){return Number(m.qty)!==0;});if(moves.length){lines.push('แจก / เทสเตอร์ / ชำรุด');moves.forEach(function(m){lines.push('• '+lbKindName_(m.kind)+' '+String(m.name).slice(0,60)+' '+m.qty+' ชิ้น');});}
 (d.counts||[]).forEach(function(c){var sku=c.sku||String(c.key).split('|')[0];if(!balance[sku])balance[sku]={expected:0,actual:0};balance[sku].expected+=Number(c.expected);balance[sku].actual+=Number(c.actual);});
 lines.push('คงเหลือ: ตามระบบ / นับจริง');Object.keys(balance).forEach(function(sku){lines.push('• '+(names[sku]||sku)+' '+balance[sku].expected+' / '+balance[sku].actual+' ชิ้น');});
 if(d.mismatch)lines.push('⚠️ มีส่วนต่าง: '+String(d.note||'').slice(0,200));
 return lines.join('\n');
}
function lbDeliver_(id){
 var sh=lbSetup_().getSheetByName(LB_LEDGER),rows=posRows_(sh),idx=rows.findIndex(function(r){return String(r[0])===id;});
 if(idx<0||String(rows[idx][8]).indexOf('PENDING')!==0)return;
 // Only new marked bill notifications or CLOSE entries may be sent.
 var e=lbEvents_().filter(function(e){return e.id===id;})[0];
 if(!lbShouldNotify_(e))return;
 try{
  // Persist immutable message before the first attempt: retries must use identical bodies.
  var payload=e.data,legacyAttempt=!!e.data.notificationText;
  if(!payload.notificationText){payload.notificationText=lbNotifyText_(e)+lbReceiptLink_(e);var serialized=JSON.stringify(payload);if(serialized.length>49000){payload.notificationText='บูธ '+lbState_(e.issue).name+' วันที่ '+e.day+'\nรายการสะสมยาวมาก กรุณาดูรายการทั้งหมดในหน้า POS\nยอดขายสะสม '+lbSummary_(lbState_(e.issue),e.day).total.toFixed(2)+' บาท'+lbReceiptLink_(e);serialized=JSON.stringify(payload);}sh.getRange(idx+2,8).setValue(serialized);SpreadsheetApp.flush();}
  var msg=payload.notificationText,chunks=[];
  while(msg.length){var part=msg.slice(0,4200);chunks.push(part);msg=msg.slice(part.length);}
  var batchCount=Math.ceil(chunks.length/5),done=lbLineCompletedBatches_(payload,rows[idx][8],batchCount);
  // Fully accepted messages can be reconciled without another LINE request.
  var selected=done<batchCount?lbSelectLineBot_():null,group=selected&&selected.groupId,token=selected&&selected.token;
  for(var batch=done;batch<batchCount;batch++){
   var uuid=e.id.replace(/^[A-Z]+-/, ''),retryKey=uuid.slice(0,24)+('000000000000'+batch.toString(16)).slice(-12);
   // Record before network I/O. A crash/timeout must be treated as uncertain.
   var prior=payload.lineAttempt;
   lbGuardLineAttempt_(prior,selected,batch,rows[idx][9],legacyAttempt&&!prior);
   var same=prior&&prior.batch===batch&&prior.botId===selected.botId&&prior.groupId===group;
   payload.lineAttempt={botId:selected.botId,basicId:selected.basicId,slot:selected.slot,groupId:group,batch:batch,firstAt:same?prior.firstAt:new Date().toISOString(),outcome:'UNCERTAIN'};
   sh.getRange(idx+2,8).setValue(JSON.stringify(payload));SpreadsheetApp.flush();
   var res=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+token,'X-Line-Retry-Key':retryKey},payload:JSON.stringify({to:group,messages:chunks.slice(batch*5,batch*5+5).map(function(t){return {type:'text',text:t};})}),muteHttpExceptions:true});
   var code=res.getResponseCode(),headers=typeof res.getAllHeaders==='function'?res.getAllHeaders():{},accepted409=code===409&&Object.keys(headers).some(function(k){return k.toLowerCase()==='x-line-accepted-request-id'&&!!headers[k];});
   if((code<200||code>=300)&&!accepted409){
    // Definite rejection allows a later bot change; 5xx and unverified 409 do not.
    payload.lineAttempt.outcome=(code>=400&&code<500&&code!==408&&code!==409&&!(same&&prior.outcome!=='REJECTED'))?'REJECTED':'UNCERTAIN';
    payload.lineAttempt.http=code;
    sh.getRange(idx+2,8).setValue(JSON.stringify(payload));SpreadsheetApp.flush();
    throw Error(lbLineError_(code,res.getContentText(),token));
   }
   payload.lineAttempt.outcome='ACCEPTED';payload.lineBatchesDone=batch+1;
   sh.getRange(idx+2,8).setValue(JSON.stringify(payload));SpreadsheetApp.flush();
   sh.getRange(idx+2,9).setValue('PENDING');SpreadsheetApp.flush();
  }
  sh.getRange(idx+2,9,1,2).setValues([['SENT','']]);
 }catch(err){sh.getRange(idx+2,10).setValue(String(err.message).slice(0,300));throw err;}
}
// V12.1.15: progress belongs in JSON; Notify remains compatible with sheet validation.
function lbLineCompletedBatches_(payload,status,batchCount){
 var done=Number(payload.lineBatchesDone||0),legacy=Number(String(status).split(':')[1]||0),a=payload.lineAttempt;
 if(!Number.isInteger(done)||done<0||done>batchCount)throw Error('ความคืบหน้าการส่งไม่ถูกต้อง กรุณาตรวจรายการ');
 if(Number.isInteger(legacy)&&legacy>=0&&legacy<=batchCount)done=Math.max(done,legacy);
 if(a&&a.outcome==='ACCEPTED'&&Number.isInteger(a.batch)&&a.batch>=0&&a.batch<batchCount)done=Math.max(done,a.batch+1);
 return done;
}
// Editor-run repair: only independently recorded acceptance; never sends LINE or modifies amounts.
function lbRepairAcceptedLineStatuses(){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  var sh=lbSetup_().getSheetByName(LB_LEDGER),rows=posRows_(sh),repaired=[];
  rows.forEach(function(row,i){
   if(['SUMMARY','SALE','VOID','CLOSE'].indexOf(String(row[2]))<0||String(row[8]).indexOf('PENDING')!==0)return;
   try{
    var p=JSON.parse(String(row[7]));if(!p.notificationText)return;
    var batches=Math.ceil(Math.ceil(p.notificationText.length/4200)/5);
    if(batches>0&&lbLineCompletedBatches_(p,row[8],batches)===batches){sh.getRange(i+2,9,1,2).setValues([['SENT','']]);repaired.push(String(row[0]));}
   }catch(ignore){} // uncertain rows stay pending for review, never guessed as sent
  });
  SpreadsheetApp.flush();var result={repaired:repaired.length,ids:repaired};Logger.log(JSON.stringify(result));return result;
 }finally{lock.releaseLock();}
}
// Fail closed when switching channels could duplicate a previously accepted request.
function lbGuardLineAttempt_(prior,selected,batch,legacyError,legacyAttempt){
 if(!prior){
  if(legacyAttempt&&!/LINE HTTP 4(?!08|09)\d\d\b/.test(String(legacyError||'')))throw Error('พักการแจ้ง: ข้อความจากรุ่นเดิมอาจส่งแล้ว กรุณาตรวจในกลุ่มก่อน ไม่ต้องลงบิลซ้ำ');
  return;
 }
 if(prior.batch!==batch)return; // previous batch was committed; never replay it
 if(prior.outcome==='REJECTED')return;
 var age=Date.now()-Date.parse(prior.firstAt);
 if(prior.botId!==selected.botId||prior.groupId!==selected.groupId)throw Error('พักการแจ้ง: บอตหรือกลุ่มเปลี่ยน แต่ยังยืนยันผลส่งครั้งก่อนไม่ได้ กรุณาตรวจในกลุ่มก่อน');
 if(!isFinite(age)||age<0||age>=23*60*60*1000)throw Error('พักการแจ้ง: เกินช่วงส่งซ้ำอย่างปลอดภัย กรุณาตรวจในกลุ่มก่อน ไม่ต้องลงบิลซ้ำ');
}
// GET diagnostics only. Tokens remain server-side and are never included in results.
function lbLineCandidates_(){
 var prop=PropertiesService.getScriptProperties(),out=[],seen=[];
 ['POS_LINE_ACCESS_TOKEN','POS_LINE_ACCESS_TOKEN_1','POS_LINE_ACCESS_TOKEN_2'].forEach(function(slot){
  var token=String(prop.getProperty(slot)||'').trim();
  if(slot==='POS_LINE_ACCESS_TOKEN'&&!token)token=String(prop.getProperty('ACCESS_TOKEN')||'').trim();
  if(token&&seen.indexOf(token)<0){out.push({slot:slot,token:token});seen.push(token);}
 });
 return out;
}
function lbSelectLineBot_(){
 var group=String(PropertiesService.getScriptProperties().getProperty('POS_GROUP_ID')||'').trim(),candidates=lbLineCandidates_(),checks=[],active=[];
 if(!group)throw Error('ยังไม่ได้ตั้งค่า POS_GROUP_ID');
 if(!candidates.length)throw Error('ยังไม่ได้ตั้งค่า POS_LINE_ACCESS_TOKEN / POS_LINE_ACCESS_TOKEN_1 / POS_LINE_ACCESS_TOKEN_2');
 var requests=candidates.map(function(c){return {url:'https://api.line.me/v2/bot/group/'+encodeURIComponent(group)+'/summary',method:'get',headers:{Authorization:'Bearer '+c.token},muteHttpExceptions:true};}),responses;
 try{responses=UrlFetchApp.fetchAll(requests);}catch(ignore){throw Error('ตรวจบอตในกลุ่มไม่ได้ กรุณาตรวจการเชื่อมต่อ LINE แล้วลองอีกครั้ง');}
 candidates.forEach(function(c,i){
  try{
   var response=responses[i],code=response.getResponseCode();
   if(code!==200){checks.push(c.slot+': HTTP '+code);return;}
   var summary=JSON.parse(response.getContentText());
   if(summary.groupId!==group){checks.push(c.slot+': กลุ่มไม่ตรง');return;}
   var info=UrlFetchApp.fetch('https://api.line.me/v2/bot/info',{method:'get',headers:{Authorization:'Bearer '+c.token},muteHttpExceptions:true});
   if(info.getResponseCode()!==200){checks.push(c.slot+': อ่านชื่อบอตไม่ได้');return;}
   var bot=JSON.parse(info.getContentText());if(!bot.userId)throw Error('missing bot ID');
   if(!active.some(function(b){return b.botId===bot.userId;}))active.push({token:c.token,slot:c.slot,botId:bot.userId,basicId:String(bot.basicId||''),name:String(bot.displayName||''),groupId:group,groupName:String(summary.groupName||'')});
  }catch(ignore){checks.push(c.slot+': ตรวจการเชื่อมต่อไม่ได้');}
 });
 if(!active.length)throw Error('ไม่มีบอตที่ยืนยันสิทธิ์ในกลุ่มได้ · '+checks.join(' · ')+' · ตรวจ Token และเชิญบอตเข้ากลุ่ม');
 if(active.length!==1)throw Error('พบหลายบอตที่เข้าถึงกลุ่มได้ กรุณาตรวจ POS_GROUP_ID และการตั้งค่า');
 active[0].checks=checks;return active[0];
}
/** A private Drive link for this bill only. Never widens sharing. */
function lbReceiptLink_(event){
 if(event.type!=='SALE'||!event.data.photoId)return '';
 var photo=lbEvents_().filter(function(e){return e.id===event.data.photoId&&e.type==='PHOTO'&&e.issue===event.issue&&e.data.kind==='receipt';})[0];
 if(!photo||!photo.data.fileId||!/^[-A-Za-z0-9_]+$/.test(photo.data.fileId))return '';
 return '\n📎 รูปแนบบิล '+event.id+'\nhttps://drive.google.com/file/d/'+photo.data.fileId+'/view';
}
// Stable issue number avoids duplicate folders when a booth is renamed.
function lbUploadFolder_(root,issue,day){
 if(!/^EVT-[A-Za-z0-9-]+$/.test(issue))throw Error('เลขใบเบิกบูธไม่ถูกต้อง');
 if(day!==undefined&&(!/^\d{4}-\d{2}-\d{2}$/.test(day)||isNaN(Date.parse(day+'T00:00:00Z'))||new Date(day+'T00:00:00Z').toISOString().slice(0,10)!==day))throw Error('วันที่จัดเก็บไม่ถูกต้อง');
 var folders=root.getFoldersByName(issue),booth=folders.hasNext()?folders.next():root.createFolder(issue);
 if(day===undefined)return booth;
 var dated=booth.getFoldersByName(day);return dated.hasNext()?dated.next():booth.createFolder(day);
}
function lbUpload_(p,role){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  if(p.kind==='product')lbRequireBoss_(role);
  if(!p.requestId)throw Error('ไม่พบเลขคำขอแนบรูป');
  var all=lbEvents_(),old=all.filter(function(e){return e.requestId===p.requestId;})[0];
  if(old){if(old.type!=='PHOTO'||old.issue!==(p.issueNo||''))throw Error('เลขคำขอแนบรูปซ้ำ');return {success:true,duplicate:true,id:old.id,data:old.data};}
  var prop=PropertiesService.getScriptProperties(),folderId=prop.getProperty('POS_UPLOAD_FOLDER_ID');
  if(!folderId)throw Error('ตั้งค่าโฟลเดอร์รูปก่อน');
  if(!/^image\/(jpeg|png|webp)$/.test(p.mime)||typeof p.base64!=='string'||p.base64.length>2800000)throw Error('รองรับ JPG PNG WEBP ไม่เกิน 2 MB');
  var bytes=Utilities.base64Decode(p.base64);if(bytes.length>2000000)throw Error('รูปเกิน 2 MB');
  if(p.kind!=='product'){var state=lbState_(p.issueNo,all);if(!state.received||state.closed||state.dayClosed)throw Error('บูธยังไม่เปิดขายหรือปิดยอดแล้ว');}
  var root=DriveApp.getFolderById(folderId),folder=root;
  if(p.kind!=='product')folder=lbUploadFolder_(root,p.issueNo,lbDay_());
  var file=folder.createFile(Utilities.newBlob(bytes,p.mime,'POS-'+Utilities.getUuid()+'.'+(p.mime==='image/jpeg'?'jpg':p.mime.split('/')[1])));
  var data={url:file.getUrl(),image:p.kind==='product'?file.getUrl():'',fileId:file.getId(),folderId:folder.getId(),kind:p.kind==='product'?'product':'receipt'};
  return lbAppend_('PHOTO',p,data);
 }finally{lock.releaseLock();}
}
function lbApi_(req){try{var prop=PropertiesService.getScriptProperties(),key=prop.getProperty('POS_API_KEY');if(!key||req.apiKey!==key)throw Error('ไม่ได้รับอนุญาต');var p=req.payload||{},role=req.role;if(req.action==='login'){var cache=CacheService.getScriptCache(),bucket='pos-login-'+String(req.clientKey||'unknown'),tries=Number(cache.get(bucket)||0);if(tries>=10)throw Error('ลองเข้าสู่ระบบมากเกินไป รอ 10 นาที');cache.put(bucket,String(tries+1),600);var boss=prop.getProperty('BOSS_PIN'),team=prop.getProperty('POS_TEAM_PIN');if(boss&&String(p.pin)===boss)role='boss';else if(team&&String(p.pin)===team)role='team';else throw Error('PIN ไม่ถูกต้อง');cache.remove(bucket);return lbJson_({success:true,role:role});}if(['boss','team'].indexOf(role)<0)throw Error('กรุณาเข้าสู่ระบบ');var result;if(req.action==='lineSummaryPreview'){lbRequireBoss_(role);result=lbManualSummary_(p.issueNo);}else if(req.action==='sendLineSummary'){lbRequireBoss_(role);result=lbSendManualSummary_(p);}else if(req.action==='lineQuota')result=lbLineQuota_();else if(req.action==='history')result=lbHistory_(p);else if(req.action==='productImages'){var wanted=Array.isArray(p.skus)?p.skus.slice(0,20):[];result=lbCatalog_().filter(function(x){return wanted.indexOf(x.sku)>=0;}).map(function(x){return {sku:x.sku,imageVersion:x.imageVersion,image:lbProductImage_(x.imageSource,x.sku)};});}else if(req.action==='init')result=lbSnapshot_(p.issueNo||'');else if(['receive','sale','void','closeDay'].indexOf(req.action)>=0)result=lbMutate_(req.action,p,role);else if(['restock','receiveRestock','stockOut'].indexOf(req.action)>=0)result=lbMovement_(req.action,p,role);else if(req.action==='upload'){if(p.kind==='product')lbRequireBoss_(role);result=lbUpload_(p,role);}else if(req.action==='promotionSave'){if(role!=='boss')throw Error('เฉพาะบอสเท่านั้นที่แก้โปรโมชั่นได้');result=lbPromotionSave_(p);}else if(req.action==='bundlePriceSave'){lbRequireBoss_(role);result=lbSaveBundlePrice_(p);}else if(req.action==='catalogSave'){lbRequireBoss_(role);var price=posMoney_(p.price);var url=String(p.image||'');if(url&&!/^https:\/\//.test(url))throw Error('รูปต้องเป็น HTTPS');var sh=lbSetup_().getSheetByName(POS_SHEETS_.products),rows=posRows_(sh),i=rows.findIndex(function(r){return String(r[0])===p.sku;});if(i<0)throw Error('ไม่พบสินค้า');sh.getRange(i+2,3,1,2).setValues([[price,url]]);result={success:true};}else if(req.action==='retryNotify'){lbRequireBoss_(role);var deliveryLock=LockService.getScriptLock();deliveryLock.waitLock(25000);try{var pending=lbEvents_().filter(function(e){return lbShouldNotify_(e)&&e.notify.indexOf('PENDING')===0;}).slice(0,10);pending.forEach(function(e){try{lbDeliver_(e.id);}catch(ignore){}});var after=lbEvents_(),statuses=pending.map(function(e){return after.filter(function(x){return x.id===e.id;})[0];});result={success:true,attempted:pending.length,sent:statuses.filter(function(e){return e&&e.notify==='SENT';}).length,remaining:after.filter(function(e){return lbShouldNotify_(e)&&e.notify.indexOf('PENDING')===0;}).length};}finally{deliveryLock.releaseLock();}}else throw Error('คำสั่งไม่รองรับ');return lbJson_({success:true,result:result});}catch(err){var rejected=false;if(key&&req.apiKey===key)try{rejected=!req.payload||!req.payload.requestId||!lbEvents_().some(function(e){return e.requestId===req.payload.requestId;});}catch(ignore){}return lbJson_({success:false,error:String(err.message||err),rejected:rejected});}}
function lbJson_(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}
function setupLilbullePOSV11(){lbSetup_();return 'สร้างตาราง V11 สำเร็จ ไม่แก้ไขยอดสต็อกเดิม';}

/** V11.1 operations: replenishments stay under the original EventIssue number. */
function lbKindName_(kind){return {FREE:'แจก',TESTER:'เทสเตอร์',DAMAGE:'ชำรุด'}[kind]||kind;}
function lbMovementSummary_(state,day){var map={};(state.movements||[]).filter(function(e){return e.type==='OUT'&&e.day===day;}).forEach(function(e){e.data.lines.forEach(function(l){var key=e.data.kind+'|'+l.sku;if(!map[key])map[key]={kind:e.data.kind,sku:l.sku,name:l.name,qty:0};map[key].qty+=l.qty;});});return Object.keys(map).map(function(k){return map[k];});}
function lbNonSaleLot_(issue,sku,exp){return lbEvents_().filter(function(e){return e.issue===issue&&e.type==='OUT';}).reduce(function(q,e){return q+e.data.lines.reduce(function(a,l){return a+(l.sku===sku&&normalizeExpKey(l.exp)===normalizeExpKey(exp)?l.qty:0);},0);},0);}
function lbUnreturnableLot_(issue,sku,exp){var state=lbState_(issue),l=state.lots.filter(function(l){return l.sku===sku&&normalizeExpKey(l.exp)===normalizeExpKey(exp);})[0];return l?l.sold+l.out+l.pending:0;}
function lbWarehouse_(){var ss=lbSetup_(),sh=ss.getSheetByName('Stock');if(!sh)return [];var result={};getStockRows(sh).forEach(function(r){if(r.qty<=0)return;var sku;try{sku=lbSku_(r.full);}catch(ignore){return;}var key=lbKey_(sku,r.exp);if(!result[key])result[key]={key:key,sku:sku,exp:r.exp,qty:0};result[key].qty+=r.qty;});return Object.keys(result).map(function(k){return result[k];});}
function lbCell_(v){return {userEnteredValue:typeof v==='number'?{numberValue:v}:{stringValue:String(v===undefined||v===null?'':v)}};}
function lbAppendRequest_(sheet,rows){return {appendCells:{sheetId:sheet.getSheetId(),rows:rows.map(function(r){return {values:r.map(lbCell_)};}),fields:'userEnteredValue'}};}
function lbAtomicRestock_(){throw Error('ปิดการตัดคลังผ่าน POS แล้ว กรุณาเบิกจริงในระบบคลังแยก');}
function lbMovement_(action,p,role){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  p.requestId=posText_(p.requestId,120);p.staff=posText_(p.staff,80);p.issueNo=posText_(p.issueNo,100);
  if(!/^[A-Za-z0-9_-]{1,120}$/.test(p.requestId)||!p.staff||!p.issueNo)throw Error('ข้อมูลรายการไม่ครบ');
  var types={restock:'RESTOCK',receiveRestock:'RESTOCK_RECEIVE',stockOut:'OUT'},events=lbEvents_(),old=events.filter(function(e){return e.requestId===p.requestId;})[0];
  if(old){if(old.type!==types[action]||old.issue!==p.issueNo)throw Error('รหัสคำขอซ้ำกับรายการอื่น');return {success:true,id:old.id,data:old.data,duplicate:true};}
  var state=lbState_(p.issueNo,events);if(state.closed||state.dayClosed)throw Error('บูธปิดงานหรือปิดยอดวันนี้แล้ว');if(!state.received)throw Error('ตรวจรับใบเบิกเริ่มต้นก่อน');
  if(action==='receiveRestock'){
   var transfer=state.pendingRestocks.filter(function(t){return t.id===p.transferId;})[0];if(!transfer)throw Error('ไม่พบรายการเติมที่รอตรวจรับ');
   if(!Array.isArray(p.counts)||p.counts.length!==transfer.data.lines.length)throw Error('นับให้ครบทุกล็อตที่เติม');
   var seen={};p.counts.forEach(function(c){var l=transfer.data.lines.filter(function(l){return l.key===c.key;})[0];if(!l||seen[c.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[c.key]=true;if(lbQty_(c.actual,true)!==l.qty)throw Error('จำนวนเติมไม่ตรง กรุณาให้ทีมงานตรวจใหม่');});
   return lbAppend_('RESTOCK_RECEIVE',p,{transferId:transfer.id,counts:p.counts,note:posText_(p.note,500)});
  }
  if(!Array.isArray(p.items)||!p.items.length||p.items.length>30)throw Error('เลือกสินค้า 1–30 ล็อต');
  if(posText_(p.note,500).length<5)throw Error('ระบุเหตุผลหรือหมายเหตุอย่างน้อย 5 ตัวอักษร');
  var catalog=lbCatalog_(),seen={},lines=[];
  if(action==='stockOut'){
   if(['FREE','TESTER','DAMAGE'].indexOf(p.kind)<0)throw Error('ประเภทรายการไม่ถูกต้อง');
   p.items.forEach(function(x){var l=state.lots.filter(function(l){return l.key===x.key;})[0];if(!l||seen[x.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[x.key]=true;var qty=lbQty_(x.qty,false);var product=catalog.filter(function(c){return c.sku===l.sku;})[0];lines.push({key:l.key,sku:l.sku,exp:l.exp,name:product?product.name:l.sku,qty:qty});});
   return lbAppend_('OUT',p,{kind:p.kind,lines:lines,note:posText_(p.note,500)});
  }
  if(action!=='restock')throw Error('คำสั่งไม่รองรับ');lbRequireBoss_(role);
  p.items.forEach(function(x){var sku=posText_(x.sku,60),exp=formatExpForSheet_(x.exp),key=lbKey_(sku,exp);if(seen[key])throw Error('เลือกล็อตซ้ำ');seen[key]=true;
    if(!exp||exp==='-')throw Error('กรุณาเลือกวันหมดอายุของ '+sku);
    var qty=lbQty_(x.qty,false),product=catalog.filter(function(c){return c.sku===sku;})[0];if(!product)throw Error('ไม่พบสินค้าใน POS');
    lines.push({key:key,sku:sku,full:getProductInfo(sku).full,exp:exp,name:product.name,qty:qty});
  });
  // Record only a POS virtual transfer; warehouse staff records physical issue separately.
  // NO Stock, EventIssue, AuditLog, SalesLog or StockIn writes here.
  return lbAppend_('RESTOCK',p,{lines:lines,note:posText_(p.note,500),inventoryEffect:'NONE',requiresWarehouseIssue:true});
 }finally{lock.releaseLock();}
}
/** Read only images explicitly linked in product catalogue. No generated substitute. */
function lbProductImage_(source,sku){
 var originals={balm:'195sebGbaFBHZYsN6LrWO2hh23cItp4m2',lotion:'1Z4N0NiPeoxVNzxykMsNnUUbmh12bkYyC',foam:'1G0KCH4GrfC6U9gAZc9M-blJytatW0PIs',mozzie:'1FuZNXwJxdrZBJRaFTio6iYTPPPLW30li',spray:'1qpFstzFhzdkvcfg4pmAmA951x72hKzUz'};
 if(!source&&originals[sku])source='https://drive.google.com/file/d/'+originals[sku]+'/view';
 if(!source)return '';var match=String(source).match(/\/d\/([A-Za-z0-9_-]+)/)||String(source).match(/[?&]id=([A-Za-z0-9_-]+)/);
 if(!match)return /^https:\/\//.test(source)?source:'';
 try{var cache=CacheService.getScriptCache(),key='lb-product-image-'+match[1],saved=cache.get(key);if(saved)return saved;
  var file=DriveApp.getFileById(match[1]),mime=file.getMimeType();if(!/^image\/(jpeg|png|webp)$/.test(mime))return '';
  var blob=file.getThumbnail();if(!blob&&file.getSize()<=350000)blob=file.getBlob();if(!blob)return '';
  var result='data:'+blob.getContentType()+';base64,'+Utilities.base64Encode(blob.getBytes());if(result.length<90000)cache.put(key,result,600);return result;
 }catch(ignore){return '';}
}

function lbHistoryRows_(events,booths,p){
 var voids={},photos={};events.forEach(function(e){if(e.type==='VOID')voids[e.issue+'|'+e.data.billId]=e;if(e.type==='PHOTO'&&e.data.kind==='receipt')photos[e.id]=e;});
 var query=String(p.query||'').trim().toLowerCase(),day=String(p.day||'');
 if(day&&!/^\d{4}-\d{2}-\d{2}$/.test(day))throw Error('วันที่ไม่ถูกต้อง');
 var rows=events.filter(function(e){return e.type==='SALE'&&(!p.issueNo||e.issue===p.issueNo)&&(!day||e.day===day);}).map(function(e){
 var v=voids[e.issue+'|'+e.id],photo=photos[e.data.photoId],b=booths.filter(function(b){return b.id===e.issue;})[0];
 return {id:e.id,issue:e.issue,boothName:b?b.name:e.issue,day:e.day,at:e.at,staff:e.staff,payment:e.data.payment,total:e.data.total,lines:e.data.lines,note:e.data.note||'',status:v?'VOID':'PAID',voidNote:v?v.data.note:'',voidStaff:v?v.staff:'',voidAt:v?v.at:'',notify:e.notify,photoUrl:photo&&photo.issue===e.issue&&/^[-A-Za-z0-9_]+$/.test(photo.data.fileId||'')?'https://drive.google.com/file/d/'+photo.data.fileId+'/view':''};
 }).filter(function(r){return (!p.status||r.status===p.status)&&(!query||[r.id,r.staff,r.note,r.boothName,r.payment].concat(r.lines.map(function(l){return l.name;})).join(' ').toLowerCase().indexOf(query)>=0);}).sort(function(a,b){return b.at.localeCompare(a.at)||b.id.localeCompare(a.id);});
 var offset=Math.max(0,Math.floor(Number(p.offset)||0)),active=rows.filter(function(r){return r.status==='PAID';});
 return {rows:rows.slice(offset,offset+50),count:rows.length,offset:offset,total:active.reduce(function(a,r){return a+r.total;},0),paidCount:active.length,voidCount:rows.length-active.length,booths:booths};
}
function lbHistory_(p){var booths={};posRows_(lbSetup_().getSheetByName('EventIssue')).forEach(function(r){booths[String(r[1])]={id:String(r[1]),name:String(r[2])};});return lbHistoryRows_(lbEvents_(),Object.keys(booths).map(function(k){return booths[k];}),p);}

// V12.1.6 | LINE quota diagnostics: GET only, never sends or retries notifications.
function lbLineSafeMessage_(raw,token){
 var message='';try{var body=JSON.parse(String(raw||''));message=typeof body.message==='string'?body.message:'';}catch(ignore){}
 if(token)message=message.split(token).join('[REDACTED]');
 return message.replace(/Bearer\s+[^\s"']+/gi,'Bearer [REDACTED]').replace(/[\r\n\t]/g,' ').slice(0,220);
}
function lbLineError_(code,raw,token){
 var message=lbLineSafeMessage_(raw,token);return 'LINE HTTP '+code+(message?' · '+message:'');
}
function lbLineFailureKind_(message){
 var text=String(message||'');
 if(/monthly limit/i.test(text))return 'MONTHLY_LIMIT';
 if(/429/.test(text)&&/rate limit|too many requests/i.test(text))return 'RATE_LIMIT';
 if(/429/.test(text))return 'UNKNOWN_429';
 if(/HTTP (401|403)/.test(text))return 'AUTH';return 'OTHER';
}
function lbLineQuotaValues_(quota,usage){
 var limited=quota&&quota.type==='limited',limit=limited&&typeof quota.value==='number'&&isFinite(quota.value)&&quota.value>=0?quota.value:null;
 var used=usage&&typeof usage.totalUsage==='number'&&isFinite(usage.totalUsage)&&usage.totalUsage>=0?usage.totalUsage:null;
 return {type:quota&&quota.type||'unknown',limit:limit,used:used,remaining:limit!==null&&used!==null?Math.max(0,limit-used):null};
}
function lbLineRead_(path,token){
 try{
  var response=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/'+path,{method:'get',headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
  var code=response.getResponseCode(),raw=response.getContentText();
  if(code!==200)return {ok:false,error:lbLineError_(code,raw,token)};
  var body;try{body=JSON.parse(raw);}catch(ignore){return {ok:false,error:'LINE ส่งข้อมูลที่อ่านไม่ได้'};}
  return {ok:true,data:body};
 }catch(ignore){return {ok:false,error:'เชื่อมต่อ LINE ไม่สำเร็จ กรุณาตรวจอีกครั้งภายหลัง'};}
}
function lbLineQuota_(){
 var prop=PropertiesService.getScriptProperties(),selected=lbSelectLineBot_(),token=selected.token;
 var quota=lbLineRead_('quota',token),usage=lbLineRead_('quota/consumption',token),values=lbLineQuotaValues_(quota.ok?quota.data:null,usage.ok?usage.data:null);
 var result={checkedAt:new Date().toISOString(),notificationPolicy:'MANUAL_SUMMARY',approximate:true,type:values.type,limit:values.limit,used:values.used,remaining:values.remaining,errors:[],pendingCount:null,failures:[]};
 result.activeBot={name:selected.name,basicId:selected.basicId,slot:selected.slot,groupName:selected.groupName};
 result.botChecks=selected.checks;
 if(!quota.ok)result.errors.push(quota.error);if(!usage.ok)result.errors.push(usage.error);
 try{
  // Read the existing ledger directly: diagnostic checks must not create sheets or modify bills.
  var id=prop.getProperty('SS_ID');if(!id)throw Error('SS_ID missing');
  var sheet=SpreadsheetApp.openById(id).getSheetByName(LB_LEDGER);
  var rows=sheet?posRows_(sheet):[],pending=rows.filter(function(row){return lbRowShouldNotify_(row)&&String(row[8]).indexOf('PENDING')===0;});
  result.pendingCount=pending.length;
  result.failures=pending.slice(-5).reverse().map(function(row){var message=String(row[9]||'').split(token).join('[REDACTED]').slice(0,300);return {id:String(row[0]),day:lbNormalizeDay_(row[4]),kind:lbLineFailureKind_(message),message:message||'ยังไม่มีรายละเอียดข้อผิดพลาด'};});
 }catch(ignore){result.errors.push('อ่านรายการข้อความค้างไม่ได้ แต่ยังแสดงผลโควตาจาก LINE ได้');}
 return result;
}
// Run from Apps Script editor for diagnosis even before updating the web. Does not send LINE.
function lbCheckLineQuota(){var result=lbLineQuota_();Logger.log(JSON.stringify(result,null,2));return result;}

// Optional cleanup affects only stale bill notification columns, never sales or new marked bills.
function lbMarkOldBillNotificationsSkipped(){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{var sh=lbSetup_().getSheetByName(LB_LEDGER),rows=posRows_(sh),skipped=0;
 rows.forEach(function(row,i){if(['SALE','VOID'].indexOf(String(row[2]))>=0&&String(row[8]).indexOf('PENDING')===0&&!lbRowShouldNotify_(row)){sh.getRange(i+2,9,1,2).setValues([['NONE','งดส่งรายบิลเก่าก่อนเปิดโหมดแจ้งทุกบิล']]);skipped++;}});
 SpreadsheetApp.flush();var result={policy:'BILL_AND_CLOSE',skippedOldBillNotifications:skipped};Logger.log(JSON.stringify(result));return result;
 }finally{lock.releaseLock();}
}

// Global suggested set prices; composition remains fixed and each actual bill retains its own price.
var LB_BUNDLE_DEFAULTS=[
 {id:'balm-lotion',name:'บาม + โลชั่น',price:690,components:[{sku:'balm',qty:1},{sku:'lotion',qty:1}]},
 {id:'balm-duo',name:'บาม 2 ชิ้น',price:590,components:[{sku:'balm',qty:2}]},
 {id:'mozzie-duo',name:'สเปรย์กันยุง 2 ชิ้น',price:350,components:[{sku:'mozzie',qty:2}]},
 {id:'foam-duo',name:'โฟม 2 ชิ้น',price:490,components:[{sku:'foam',qty:2}]},
 {id:'spray-duo',name:'แอลกอฮอล์ 2 ชิ้น',price:150,components:[{sku:'spray',qty:2}]},
 {id:'set-1',name:'ชุด SET 1 · ครบ 5 ตัว',price:1290,components:[{sku:'balm',qty:1},{sku:'lotion',qty:1},{sku:'mozzie',qty:1},{sku:'foam',qty:1},{sku:'spray',qty:1}]}
 ,{id:'spray-four',name:'แอลกอฮอล์ 4 ขวด',price:300,components:[{sku:'spray',qty:4}]}
 ,{id:'mozzie-trio',name:'กันยุง 3 ขวด',price:450,components:[{sku:'mozzie',qty:3}]}
 ,{id:'lotion-duo',name:'โลชั่น 2 หลอด',price:690,components:[{sku:'lotion',qty:2}]}
];
function lbBundlePrices_(){
 var ss=lbSetup_(),sh=ss.getSheetByName('LilbullePOS_BundlePrices'),rows=sh?posRows_(sh):[];
 return LB_BUNDLE_DEFAULTS.map(function(b){var row=rows.filter(function(r){return String(r[0])===b.id;})[0];return {id:b.id,price:row?Number(row[1]):b.price,revision:row?String(row[2]):''};});
}
function lbSaveBundlePrice_(p){
 var base=LB_BUNDLE_DEFAULTS.filter(function(b){return b.id===p.id;})[0];if(!base)throw Error('ไม่พบเซต');
 var price=Number(p.price);if(p.price===''||!isFinite(price)||price<=0||price>1000000||Math.abs(price*100-Math.round(price*100))>0.000001)throw Error('กรอกราคามากกว่า 0 ทศนิยมไม่เกิน 2 ตำแหน่ง และไม่เกิน 1,000,000 บาท');
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  var ss=lbSetup_(),sh=ss.getSheetByName('LilbullePOS_BundlePrices'),rows=sh?posRows_(sh):[],index=rows.findIndex(function(r){return String(r[0])===p.id;}),current=index>=0?String(rows[index][2]):'';
  if(typeof p.revision!=='string'||p.revision!==current)throw Error('ราคาเซตนี้ถูกแก้จากเครื่องอื่นแล้ว กรุณาเปิดหน้าราคาใหม่ก่อนบันทึก');
  if(!sh){posGetSheet_(ss,'LilbullePOS_BundlePrices',['ID','Price','Revision','UpdatedAt','Staff']);sh=ss.getSheetByName('LilbullePOS_BundlePrices');}
  var revision=Utilities.getUuid(),row=[p.id,Math.round(price*100)/100,revision,new Date(),String(p.staff||'').replace(/^=/,"'=")];
  if(index>=0)sh.getRange(index+2,1,1,5).setValues([row]);else sh.appendRow(row);
  SpreadsheetApp.flush();return {success:true,id:p.id,price:row[1],revision:revision};
 }finally{lock.releaseLock();}
}

// No sale, void or close sends LINE. Only an explicit authenticated summary request does.
function lbManualSummary_(issue,events){
 var state=lbState_(issue,events),day=lbDay_(),sum=lbSummary_(state,day),closed=state.events.filter(function(e){return e.type==='CLOSE'&&e.day===day;}).slice(-1)[0],text;
 if(closed)text=lbDailyCloseText_(closed,state.name);
 else{
  var payNames={CASH:'เงินสด',TRANSFER:'เงินโอน',CARD:'บัตร',OTHER:'อื่น ๆ'};
  text='บูธ '+String(state.name).slice(0,100)+'\nวันที่ '+day+' · สรุปยอดระหว่างวัน\nยอดขาย '+Number(sum.total).toFixed(2)+' บาท / '+sum.bills+' บิล';
  sum.products.filter(function(p){return p.qty!==0;}).forEach(function(p){text+='\n• '+p.name+' ขาย '+p.qty+' ชิ้น = '+Number(p.total).toFixed(2)+' บาท';});
  Object.keys(sum.money).filter(function(k){return sum.money[k]!==0;}).forEach(function(k){text+='\n'+(payNames[k]||k)+' ตามบิล '+Number(sum.money[k]).toFixed(2)+' บาท';});
  if(sum.movements.length){text+='\nแจก / เทสเตอร์ / ชำรุด';sum.movements.filter(function(m){return m.qty!==0;}).forEach(function(m){text+='\n• '+lbKindName_(m.kind)+' '+m.name+' '+m.qty+' ชิ้น';});}
  var balances={};state.lots.forEach(function(l){balances[l.sku]=(balances[l.sku]||0)+l.remaining;});var names={balm:'Baby Rich Balm',lotion:'Organic Lotion',foam:'Baby Foam',mozzie:'Mozzie Guard',spray:'Sanitizer Spray'};
  text+='\nคงเหลือตามระบบ (ยังไม่ได้ตรวจนับปิดวัน)';Object.keys(balances).forEach(function(k){text+='\n• '+(names[k]||k)+' '+balances[k]+' ชิ้น';});
 }
 if(text.length>35000)throw Error('สรุปยาวเกินกำหนด กรุณาตรวจรายการก่อนส่ง');
 var hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,issue+'|'+day+'|'+text,Utilities.Charset.UTF_8).map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
 return {issueNo:issue,day:day,total:sum.total,bills:sum.bills,text:text,previewKey:hash};
}
function lbSendManualSummary_(p){
 if(!/^[A-Za-z0-9_-]{1,120}$/.test(String(p.requestId||'')))throw Error('ไม่พบเลขคำขอส่งสรุป');
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  var events=lbEvents_(),old=events.filter(function(e){return e.requestId===p.requestId;})[0];
  if(old&&(old.type!=='SUMMARY'||old.issue!==p.issueNo))throw Error('เลขคำขอซ้ำกับรายการอื่น');
  if(!old){
   var preview=lbManualSummary_(p.issueNo,events);
   if(preview.previewKey!==p.previewKey)throw Error('ยอดเปลี่ยนจากหน้าตัวอย่างแล้ว กรุณาเปิดสรุปใหม่ก่อนส่ง');
   old=events.filter(function(e){return e.type==='SUMMARY'&&e.issue===p.issueNo&&e.day===preview.day&&e.data.previewKey===preview.previewKey;}).slice(-1)[0];
   if(!old){var added=lbAppend_('SUMMARY',p,{notificationPolicy:'MANUAL_SUMMARY',manualText:preview.text,previewKey:preview.previewKey});old={id:added.id,type:'SUMMARY',issue:p.issueNo,data:added.data,notify:'PENDING'};}
  }
  if(old.notify==='SENT')return {success:true,id:old.id,duplicate:true,notificationStatus:'SENT'};
  // User clicked send/retry; preserve the exact original body and retry key.
  try{lbDeliver_(old.id);return {success:true,id:old.id,notificationStatus:'SENT'};}
  catch(err){return {success:true,id:old.id,notificationStatus:'PENDING',notificationPending:true,notificationError:String(err.message||err).slice(0,300)};}
 }finally{lock.releaseLock();}
}


/** V12.3 - Admin-managed POS promotions, independent of warehouse stock. */
var LB_PROMO_SHEET='LilbullePOS_Promotions';
function lbPromoSheet_(){var ss=lbSetup_();return posGetSheet_(ss,LB_PROMO_SHEET,['ID','Name','Price','ComponentsJSON','Active','Revision','UpdatedAt','Staff']);}
function lbPromotions_(){
 var sh=lbPromoSheet_(), rows=posRows_(sh), overrides={};
 rows.forEach(function(r){
  var id=String(r[0]||''); if(!id)return;
  var lines;try{lines=JSON.parse(String(r[3]||'[]'));}catch(e){throw Error('โปรโมชั่น '+id+' มีรายการสินค้าที่อ่านไม่ได้');}
  overrides[id]={id:id,name:String(r[1]||''),price:Number(r[2]),components:lines,active:String(r[4]).toUpperCase()==='YES',revision:String(r[5]||'')};
 });
 var oldPrices=lbBundlePrices_();
 var seeds=LB_BUNDLE_DEFAULTS.map(function(b){var old=oldPrices.filter(function(x){return x.id===b.id;})[0];return {id:b.id,name:b.name,price:old?old.price:b.price,components:b.components,active:true,revision:''};});
 seeds.forEach(function(b){if(!overrides[b.id])overrides[b.id]=b;});
 return Object.keys(overrides).map(function(id){return overrides[id];}).sort(function(a,b){return a.name.localeCompare(b.name,'th');});
}
function lbPromotionSave_(p){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  p=p||{};
  var rawId=String(p.id||''),name=String(p.name||'').trim(),items=p.components,price=Number(p.price),active=p.active===true;
  if(!name||name.length>100||/[\r\n]/.test(name))throw Error('ชื่อโปรโมชั่นต้องมี 1–100 ตัวอักษร');
  if(!Number.isFinite(price)||price<=0||price>1000000||Math.abs(price*100-Math.round(price*100))>0.000001)throw Error('ราคาโปรโมชั่นไม่ถูกต้อง');
  if(!Array.isArray(items)||items.length<1||items.length>10)throw Error('ต้องเลือกสินค้า 1–10 รายการ');
  var catalog=lbCatalog_(),comb={},clean=[];
  items.forEach(function(it){
   var sku=String(it.sku||''),q=Number(it.qty);
   if(!catalog.some(function(c){return c.sku===sku;}))throw Error('ไม่พบสินค้า '+sku);
   if(!Number.isSafeInteger(q)||q<1||q>999)throw Error('จำนวนสินค้าในชุดไม่ถูกต้อง');
   if(comb[sku])throw Error('SKU ซ้ำในโปรโมชั่น กรุณารวมจำนวนก่อนบันทึก');comb[sku]=true;
   clean.push({sku:sku,qty:q});
  });
  if(rawId&&!/^[a-zA-Z0-9_-]{1,100}$/.test(rawId))throw Error('รหัสโปรโมชั่นไม่ถูกต้อง');
  var sh=lbPromoSheet_(),rows=posRows_(sh),idx=rows.findIndex(function(r){return String(r[0])===rawId;}),existing=lbPromotions_().filter(function(b){return b.id===rawId;})[0];
  if(rawId&&!existing)throw Error('ไม่พบโปรโมชั่นที่จะแก้');
  var current=idx>=0?String(rows[idx][5]||''):'';
  if(rawId && String(p.revision===undefined?'':p.revision)!==current)throw Error('มีคนแก้โปรโมชั่นจากเครื่องอื่นแล้ว กรุณารีเฟรชก่อนบันทึก');
  var id=rawId||('promo-'+Utilities.getUuid().slice(0,12)),revision=Utilities.getUuid();
  var row=[id,name,Math.round(price*100)/100,JSON.stringify(clean),active?'YES':'NO',revision,new Date(),String(p.staff||'').replace(/^=/,"'=")];
  if(idx>=0)sh.getRange(idx+2,1,1,8).setValues([row]);else sh.appendRow(row);
  SpreadsheetApp.flush();return {success:true,id:id,revision:revision,promotions:lbPromotions_()};
 }finally{lock.releaseLock();}
}
