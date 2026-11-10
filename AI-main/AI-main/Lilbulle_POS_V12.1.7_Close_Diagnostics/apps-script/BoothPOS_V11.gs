/** lilbulle POS V12.1.6 | 2026-10-26 | Hide empty LINE summary rows; expose release version.
 * Replace this extension in the SAME Apps Script project. API schema remains 11.4.1. */
var LB_APP_VERSION='12.1.6';
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
 selected.forEach(function(r){var sku=lbSku_(r[3]),exp=formatExpForSheet_(r[4]),key=lbKey_(sku,exp);if(!lots[key])lots[key]={key:key,sku:sku,exp:exp,issued:0,returned:0,sold:0,out:0,pending:0,remaining:0};lots[key].issued+=Number(r[5])||0;});
 var ret=ss.getSheetByName('EventReturn');if(ret)posRows_(ret).forEach(function(r){if(String(r[2])!==issue)return;var key=lbKey_(lbSku_(r[4]),r[5]);if(lots[key])lots[key].returned+=Number(r[6])||0;});
 es.filter(function(e){return e.type==='VOID';}).forEach(function(e){voids[e.data.billId]=true;});
 var sales=es.filter(function(e){return e.type==='SALE'&&!voids[e.id];});
 sales.forEach(function(e){e.data.lines.forEach(function(l){if(!lots[l.key])throw Error('ล็อตขายไม่ตรงใบเบิก');lots[l.key].sold+=l.qty;});});
 var receivedRestocks={};es.filter(function(e){return e.type==='RESTOCK_RECEIVE';}).forEach(function(e){receivedRestocks[e.data.transferId]=true;});
 var transfers=es.filter(function(e){return e.type==='RESTOCK';});
 transfers.forEach(function(e){if(!receivedRestocks[e.id])e.data.lines.forEach(function(l){if(lots[l.key])lots[l.key].pending+=l.qty;});});
 es.filter(function(e){return e.type==='OUT';}).forEach(function(e){e.data.lines.forEach(function(l){if(lots[l.key])lots[l.key].out+=l.qty;});});
 Object.keys(lots).forEach(function(k){var l=lots[k];l.remaining=l.issued-l.returned-l.sold-l.out-l.pending;});
 var receives=es.filter(function(e){return e.type==='RECEIVE';});
 return {id:issue,name:String(selected[0][2]),location:String(selected[0][8]||''),closed:selected.some(function(r){return String(r[7])==='ปิดแล้ว';}),received:receives.length>0,pendingRestocks:transfers.filter(function(e){return !receivedRestocks[e.id];}),movements:es.filter(function(e){return e.type==='OUT'||e.type==='RESTOCK'||e.type==='RESTOCK_RECEIVE';}),lots:Object.keys(lots).map(function(k){return lots[k];}),sales:sales,events:es,day:lbDay_(),dayClosed:es.some(function(e){return e.type==='CLOSE'&&e.day===lbDay_();})};
}
function lbSoldLot_(issue,sku,exp){var events=lbEvents_().filter(function(e){return e.issue===issue;}),voids={},qty=0;events.filter(function(e){return e.type==='VOID';}).forEach(function(e){voids[e.data.billId]=true;});events.filter(function(e){return e.type==='SALE'&&!voids[e.id];}).forEach(function(e){e.data.lines.forEach(function(l){if(l.sku===sku&&normalizeExpKey(l.exp)===normalizeExpKey(exp))qty+=l.qty;});});return qty;}
function lbSummary_(state,day){var list=state.sales.filter(function(e){return e.day===day;}),money={CASH:0,TRANSFER:0,CARD:0,OTHER:0},products={};list.forEach(function(e){money[e.data.payment]+=e.data.total;e.data.lines.forEach(function(l){if(!products[l.sku])products[l.sku]={sku:l.sku,name:l.name,qty:0,total:0};products[l.sku].qty+=l.qty;products[l.sku].total+=l.qty*l.price;});});Object.keys(money).forEach(function(k){money[k]=Math.round(money[k]*100)/100;});return {day:day,bills:list.length,money:money,total:Math.round(Object.keys(money).reduce(function(a,k){return a+money[k];},0)*100)/100,products:Object.keys(products).map(function(k){return products[k];}),movements:lbMovementSummary_(state,day),sales:list};}
function lbSnapshot_(issue){var events=lbEvents_(),issues={},ss=lbSetup_(),sh=ss.getSheetByName('EventIssue');if(sh)posRows_(sh).forEach(function(r){if(String(r[7])!=='ปิดแล้ว')issues[String(r[1])]={id:String(r[1]),name:String(r[2]),location:String(r[8]||'')};});var state=issue?lbState_(issue,events):null;if(state){state.closes=state.events.filter(function(e){return e.type==='CLOSE';}).map(function(e){delete e.data.notificationText;return {day:e.day,staff:e.staff,data:e.data};});delete state.events;state.sales.forEach(function(e){delete e.data.notificationText;});}return {version:'11.4.1',appVersion:LB_APP_VERSION,warehouse:lbWarehouse_(),products:lbCatalog_(),booths:Object.keys(issues).map(function(k){return issues[k];}).reverse(),state:state,summary:state?lbSummary_(state,lbDay_()):null,pendingNotifications:events.filter(function(e){return e.notify.indexOf('PENDING')===0&&(!issue||e.issue===issue);}).length};}
function lbAppend_(type,p,data){var json=JSON.stringify(data);if(json.length>44000)throw Error('ข้อมูลรายการยาวเกินกำหนด');var sh=lbSetup_().getSheetByName(LB_LEDGER),id=type+'-'+Utilities.getUuid();sh.appendRow([id,p.requestId,type,p.issueNo||'',lbDay_(),new Date(),(/^=/.test(p.staff)?"'"+p.staff:p.staff),json,(type==='SALE'||type==='VOID'||type==='CLOSE')?'PENDING':'NONE','']);SpreadsheetApp.flush();return {success:true,id:id,data:data};}
// Owner-approved policy: both authenticated POS roles have full operational access.
function lbRequireBoss_(role){if(['boss','team'].indexOf(role)<0)throw Error('กรุณาเข้าสู่ระบบ');}
function lbMutate_(action,p,role){
 var lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  p.requestId=posText_(p.requestId,120);p.staff=posText_(p.staff,80);p.issueNo=posText_(p.issueNo,100);
  if(!/^[A-Za-z0-9_-]{1,120}$/.test(p.requestId)||!p.staff)throw Error('กรอกพนักงานและรหัสคำขอ');
  var all=lbEvents_(),old=all.filter(function(e){return e.requestId===p.requestId;})[0];
  if(old){if(old.issue!==String(p.issueNo||'')||old.type!==({receive:'RECEIVE',sale:'SALE',void:'VOID',closeDay:'CLOSE'})[action])throw Error('รหัสคำขอซ้ำกับรายการอื่น');return {success:true,id:old.id,duplicate:true,data:old.data};}
  var s=lbState_(p.issueNo,all);if(s.closed)throw Error('ใบเบิกนี้ปิดงานแล้ว');
  var result;
  if(action==='receive'){
   if(s.received)throw Error('ตรวจรับบูธนี้แล้ว');
   // Existing POS sales cannot be inferred or silently migrated.
   var legacySS=lbSetup_(),legacyBooths=posRows_(legacySS.getSheetByName(POS_SHEETS_.booths)).filter(function(b){return String(b[7])===p.issueNo;}).map(function(b){return String(b[0]);});
   if(posRows_(legacySS.getSheetByName(POS_SHEETS_.bills)).some(function(b){return legacyBooths.indexOf(String(b[2]))>=0&&String(b[8])==='PAID';}))throw Error('ใบเบิกนี้มีบิล POS รุ่นเก่า ต้องตรวจย้ายยอดก่อนใช้งาน');
   if(!Array.isArray(p.counts)||p.counts.length!==s.lots.length)throw Error('ตรวจนับให้ครบทุกล็อต');
   var seen={};p.counts.forEach(function(x){var l=s.lots.filter(function(l){return l.key===x.key;})[0];if(!l||seen[x.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[x.key]=true;if(lbQty_(x.actual,true)!==l.remaining)throw Error('ยอดตรวจรับไม่ตรงใบเบิก ให้แอดมินแก้ไขเบิก/คืนก่อนเปิดขาย');});
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
     if(remaining)throw Error('สินค้าบูธไม่พอ: '+product.name);total+=Math.round(qty*price*100)/100;
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
  try{lbDeliver_(result.id);}catch(err){result.notificationPending=true;}
  return result;
 }finally{lock.releaseLock();}
}
function lbNotifyText_(event){var s=lbState_(event.issue),sum=lbSummary_(s,event.day),head='บูธ '+s.name+' วันที่ '+event.day+'\n';if(event.type==='CLOSE'){var data=event.data;return head+'📊 ปิดยอดรายวัน\n'+data.summary.products.filter(function(x){return Number(x.qty)!==0||Number(x.total)!==0;}).map(function(x){return '• '+x.name+' '+x.qty+' ชิ้น = '+x.total.toFixed(2)+' บาท';}).join('\n')+'\nยอดขาย '+data.summary.total.toFixed(2)+' บาท / '+data.summary.bills+' บิล\n'+Object.keys(data.actualMoney).filter(function(k){return Number(data.actualMoney[k])!==0||Number(data.summary.money[k]||0)!==0;}).map(function(k){return ({CASH:'เงินสด',TRANSFER:'โอนเงิน',CARD:'บัตร',OTHER:'อื่น ๆ'}[k]||k)+': รับจริง '+data.actualMoney[k].toFixed(2)+' บาท (ตามบิล '+data.summary.money[k].toFixed(2)+')';}).join('\n')+((data.summary.movements||[]).some(function(m){return Number(m.qty)!==0;})?'\nรายการไม่ใช่ยอดขาย: '+(data.summary.movements||[]).filter(function(m){return Number(m.qty)!==0;}).map(function(m){return lbKindName_(m.kind)+' '+m.name+' '+m.qty+' ชิ้น';}).join(', '):'')+'\nคงเหลือตามระบบ '+data.counts.reduce(function(a,x){return a+x.expected;},0)+' ชิ้น / นับจริง '+data.counts.reduce(function(a,x){return a+x.actual;},0)+' ชิ้น'+(data.mismatch?'\n⚠️ มีส่วนต่าง: '+data.note:'');}return head+(event.type==='VOID'?'ยกเลิกบิล '+event.data.billId+'\n':'')+sum.sales.map(function(e,i){return (i+1)+'. '+e.data.lines.map(function(l){return l.name+' ×'+l.qty+' @'+l.price.toFixed(2);}).join(', ')+' = '+e.data.total.toFixed(2)+' บาท';}).join('\n')+'\nยอดขายสะสม '+sum.total.toFixed(2)+' บาท';}
function lbDeliver_(id){
 var sh=lbSetup_().getSheetByName(LB_LEDGER),rows=posRows_(sh),idx=rows.findIndex(function(r){return String(r[0])===id;});
 if(idx<0||String(rows[idx][8]).indexOf('PENDING')!==0)return;
 var e=lbEvents_().filter(function(e){return e.id===id;})[0];
 try{
  var prop=PropertiesService.getScriptProperties(),group=prop.getProperty('POS_GROUP_ID'),token=prop.getProperty('POS_LINE_ACCESS_TOKEN')||prop.getProperty('ACCESS_TOKEN');
  if(!group||!token)throw Error('ยังไม่ได้ตั้งค่ากลุ่ม POS');
  // Persist immutable message before the first attempt: retries must use identical bodies.
  var payload=e.data;
  if(!payload.notificationText){payload.notificationText=lbNotifyText_(e)+lbReceiptLink_(e);var serialized=JSON.stringify(payload);if(serialized.length>49000){payload.notificationText='บูธ '+lbState_(e.issue).name+' วันที่ '+e.day+'\nรายการสะสมยาวมาก กรุณาดูรายการทั้งหมดในหน้า POS\nยอดขายสะสม '+lbSummary_(lbState_(e.issue),e.day).total.toFixed(2)+' บาท'+lbReceiptLink_(e);serialized=JSON.stringify(payload);}sh.getRange(idx+2,8).setValue(serialized);SpreadsheetApp.flush();}
  var msg=payload.notificationText,chunks=[];
  while(msg.length){var part=msg.slice(0,4200);chunks.push(part);msg=msg.slice(part.length);}
  var done=Number(String(rows[idx][8]).split(':')[1]||0),batchCount=Math.ceil(chunks.length/5);
  for(var batch=done;batch<batchCount;batch++){
   var uuid=e.id.replace(/^[A-Z]+-/, ''),retryKey=uuid.slice(0,24)+('000000000000'+batch.toString(16)).slice(-12);
   var res=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+token,'X-Line-Retry-Key':retryKey},payload:JSON.stringify({to:group,messages:chunks.slice(batch*5,batch*5+5).map(function(t){return {type:'text',text:t};})}),muteHttpExceptions:true});
   var code=res.getResponseCode();if((code<200||code>=300)&&code!==409)throw Error(lbLineError_(code,res.getContentText(),token));
   sh.getRange(idx+2,9).setValue('PENDING:'+(batch+1));SpreadsheetApp.flush();
  }
  sh.getRange(idx+2,9,1,2).setValues([['SENT','']]);
 }catch(err){sh.getRange(idx+2,10).setValue(String(err.message).slice(0,300));throw err;}
}
/** A private Drive link for this bill only. Never widens sharing. */
function lbReceiptLink_(event){
 if(event.type!=='SALE'||!event.data.photoId)return '';
 var photo=lbEvents_().filter(function(e){return e.id===event.data.photoId&&e.type==='PHOTO'&&e.issue===event.issue&&e.data.kind==='receipt';})[0];
 if(!photo||!photo.data.fileId||!/^[-A-Za-z0-9_]+$/.test(photo.data.fileId))return '';
 return '\n📎 รูปแนบบิล '+event.id+'\nhttps://drive.google.com/file/d/'+photo.data.fileId+'/view';
}
// Stable issue number avoids duplicate folders when a booth is renamed.
function lbUploadFolder_(root,issue){
 if(!/^EVT-[A-Za-z0-9-]+$/.test(issue))throw Error('เลขใบเบิกบูธไม่ถูกต้อง');
 var folders=root.getFoldersByName(issue);
 return folders.hasNext()?folders.next():root.createFolder(issue);
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
  if(p.kind!=='product')folder=lbUploadFolder_(root,p.issueNo);
  var file=folder.createFile(Utilities.newBlob(bytes,p.mime,'POS-'+Utilities.getUuid()+'.'+(p.mime==='image/jpeg'?'jpg':p.mime.split('/')[1])));
  var data={url:file.getUrl(),image:p.kind==='product'?file.getUrl():'',fileId:file.getId(),folderId:folder.getId(),kind:p.kind==='product'?'product':'receipt'};
  return lbAppend_('PHOTO',p,data);
 }finally{lock.releaseLock();}
}
function lbApi_(req){try{var prop=PropertiesService.getScriptProperties(),key=prop.getProperty('POS_API_KEY');if(!key||req.apiKey!==key)throw Error('ไม่ได้รับอนุญาต');var p=req.payload||{},role=req.role;if(req.action==='login'){var cache=CacheService.getScriptCache(),bucket='pos-login-'+String(req.clientKey||'unknown'),tries=Number(cache.get(bucket)||0);if(tries>=10)throw Error('ลองเข้าสู่ระบบมากเกินไป รอ 10 นาที');cache.put(bucket,String(tries+1),600);var boss=prop.getProperty('BOSS_PIN'),team=prop.getProperty('POS_TEAM_PIN');if(boss&&String(p.pin)===boss)role='boss';else if(team&&String(p.pin)===team)role='team';else throw Error('PIN ไม่ถูกต้อง');cache.remove(bucket);return lbJson_({success:true,role:role});}if(['boss','team'].indexOf(role)<0)throw Error('กรุณาเข้าสู่ระบบ');var result;if(req.action==='lineQuota')result=lbLineQuota_();else if(req.action==='history')result=lbHistory_(p);else if(req.action==='productImages'){var wanted=Array.isArray(p.skus)?p.skus.slice(0,20):[];result=lbCatalog_().filter(function(x){return wanted.indexOf(x.sku)>=0;}).map(function(x){return {sku:x.sku,imageVersion:x.imageVersion,image:lbProductImage_(x.imageSource,x.sku)};});}else if(req.action==='init')result=lbSnapshot_(p.issueNo||'');else if(['receive','sale','void','closeDay'].indexOf(req.action)>=0)result=lbMutate_(req.action,p,role);else if(['restock','receiveRestock','stockOut'].indexOf(req.action)>=0)result=lbMovement_(req.action,p,role);else if(req.action==='upload'){if(p.kind==='product')lbRequireBoss_(role);result=lbUpload_(p,role);}else if(req.action==='catalogSave'){lbRequireBoss_(role);var price=posMoney_(p.price);var url=String(p.image||'');if(url&&!/^https:\/\//.test(url))throw Error('รูปต้องเป็น HTTPS');var sh=lbSetup_().getSheetByName(POS_SHEETS_.products),rows=posRows_(sh),i=rows.findIndex(function(r){return String(r[0])===p.sku;});if(i<0)throw Error('ไม่พบสินค้า');sh.getRange(i+2,3,1,2).setValues([[price,url]]);result={success:true};}else if(req.action==='retryNotify'){lbRequireBoss_(role);var deliveryLock=LockService.getScriptLock();deliveryLock.waitLock(25000);try{var pending=lbEvents_().filter(function(e){return e.notify.indexOf('PENDING')===0;}).slice(0,10);pending.forEach(function(e){try{lbDeliver_(e.id);}catch(ignore){}});result={success:true};}finally{deliveryLock.releaseLock();}}else throw Error('คำสั่งไม่รองรับ');return lbJson_({success:true,result:result});}catch(err){var rejected=false;if(key&&req.apiKey===key)try{rejected=!req.payload||!req.payload.requestId||!lbEvents_().some(function(e){return e.requestId===req.payload.requestId;});}catch(ignore){}return lbJson_({success:false,error:String(err.message||err),rejected:rejected});}}
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
function lbAtomicRestock_(p,state,lines,stockRows,id){
 if(typeof Sheets==='undefined'||!Sheets.Spreadsheets)throw Error('กรุณาเพิ่มบริการ Google Sheets API ใน Apps Script ก่อนใช้เติมสินค้า');
 var ss=lbSetup_(),stock=ss.getSheetByName('Stock'),issue=ss.getSheetByName('EventIssue'),audit=ensureControlSheets_().auditSheet,ledger=ss.getSheetByName(LB_LEDGER),day=lbDay_(),at=nowDateTime();
 var data={lines:lines,note:posText_(p.note,500),documentNo:id};
 var requests=stockRows.filter(function(r){return r._before!==undefined&&r._before!==r.qty;}).map(function(r){return {updateCells:{range:{sheetId:stock.getSheetId(),startRowIndex:r.sheetRow-1,endRowIndex:r.sheetRow,startColumnIndex:2,endColumnIndex:3},rows:[{values:[lbCell_(r.qty)]}],fields:'userEnteredValue'}};});
 requests.push(lbAppendRequest_(issue,lines.map(function(l){return [at,state.id,state.name,l.full,l.exp,l.qty,p.staff,'เปิด',state.location];})));
 requests.push(lbAppendRequest_(audit,lines.map(function(l){return [at,id,'เติมสินค้าบูธ',l.full,l.exp,-l.qty,l.centralBefore,l.centralAfter,p.staff,'POS | อ้างอิง '+state.id,p.note];})));
 requests.push(lbAppendRequest_(ledger,[[id,p.requestId,'RESTOCK',state.id,day,at,p.staff,JSON.stringify(data),'NONE','']]));
 // A single Sheets API batch commits central stock + issue + audit + transfer together.
 SpreadsheetApp.flush();Sheets.Spreadsheets.batchUpdate({requests:requests},SS_ID);SpreadsheetApp.flush();
 return {success:true,id:id,data:data};
}
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
   var seen={};p.counts.forEach(function(c){var l=transfer.data.lines.filter(function(l){return l.key===c.key;})[0];if(!l||seen[c.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[c.key]=true;if(lbQty_(c.actual,true)!==l.qty)throw Error('จำนวนเติมไม่ตรง กรุณาให้แอดมินตรวจสอบก่อนรับ');});
   return lbAppend_('RESTOCK_RECEIVE',p,{transferId:transfer.id,counts:p.counts,note:posText_(p.note,500)});
  }
  if(!Array.isArray(p.items)||p.items.length<1||p.items.length>30)throw Error('เลือกสินค้า 1–30 ล็อต');
  if(posText_(p.note,500).length<5)throw Error('ระบุเหตุผลหรือหมายเหตุอย่างน้อย 5 ตัวอักษร');
  var catalog=lbCatalog_(),seen={},lines=[];
  if(action==='stockOut'){
   if(['FREE','TESTER','DAMAGE'].indexOf(p.kind)<0)throw Error('ประเภทรายการไม่ถูกต้อง');
   p.items.forEach(function(x){var l=state.lots.filter(function(l){return l.key===x.key;})[0];if(!l||seen[x.key])throw Error('ล็อตไม่ถูกต้องหรือซ้ำ');seen[x.key]=true;var qty=lbQty_(x.qty,false);if(qty>l.remaining)throw Error('สินค้าในบูธไม่พอ');var product=catalog.filter(function(c){return c.sku===l.sku;})[0];lines.push({key:l.key,sku:l.sku,exp:l.exp,name:product?product.name:l.sku,qty:qty});});
   return lbAppend_('OUT',p,{kind:p.kind,lines:lines,note:posText_(p.note,500)});
  }
  if(action!=='restock')throw Error('คำสั่งไม่รองรับ');lbRequireBoss_(role);
  var ss=lbSetup_(),stockRows=getStockRows(getSheetOrThrow(ss,'Stock'));
  p.items.forEach(function(x){var sku=posText_(x.sku,60),exp=formatExpForSheet_(x.exp),key=lbKey_(sku,exp);if(seen[key])throw Error('เลือกล็อตซ้ำ');seen[key]=true;var info=getProductInfo(sku),qty=lbQty_(x.qty,false),remain=qty;
   var matches=stockRows.filter(function(r){return normalizeData(r.full)===normalizeData(info.full)&&normalizeExpKey(r.exp)===normalizeExpKey(exp);}),before=matches.reduce(function(a,r){return a+r.qty;},0);
   if(before<qty)throw Error('คลังกลางไม่พอ: '+info.short+' Exp '+exp);
   matches.forEach(function(r){var cut=Math.min(remain,r.qty);if(cut>0){if(r._before===undefined)r._before=r.qty;r.qty-=cut;remain-=cut;}});
   var product=catalog.filter(function(c){return c.sku===sku;})[0];lines.push({key:key,sku:sku,full:info.full,exp:exp,name:product?product.name:sku,qty:qty,centralBefore:before,centralAfter:before-qty});
  });
  return lbAtomicRestock_(p,state,lines,stockRows,'RESTOCK-'+Utilities.getUuid());
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
 var prop=PropertiesService.getScriptProperties(),token=prop.getProperty('POS_LINE_ACCESS_TOKEN')||prop.getProperty('ACCESS_TOKEN');
 if(!token)throw Error('ยังไม่ได้ตั้งค่า POS_LINE_ACCESS_TOKEN หรือ ACCESS_TOKEN');
 var quota=lbLineRead_('quota',token),usage=lbLineRead_('quota/consumption',token),values=lbLineQuotaValues_(quota.ok?quota.data:null,usage.ok?usage.data:null);
 var result={checkedAt:new Date().toISOString(),approximate:true,type:values.type,limit:values.limit,used:values.used,remaining:values.remaining,errors:[],pendingCount:null,failures:[]};
 if(!quota.ok)result.errors.push(quota.error);if(!usage.ok)result.errors.push(usage.error);
 try{
  // Read the existing ledger directly: diagnostic checks must not create sheets or modify bills.
  var id=prop.getProperty('SS_ID');if(!id)throw Error('SS_ID missing');
  var sheet=SpreadsheetApp.openById(id).getSheetByName(LB_LEDGER);
  var rows=sheet?posRows_(sheet):[],pending=rows.filter(function(row){return String(row[8]).indexOf('PENDING')===0;});
  result.pendingCount=pending.length;
  result.failures=pending.slice(-5).reverse().map(function(row){var message=String(row[9]||'').split(token).join('[REDACTED]').slice(0,300);return {id:String(row[0]),day:lbNormalizeDay_(row[4]),kind:lbLineFailureKind_(message),message:message||'ยังไม่มีรายละเอียดข้อผิดพลาด'};});
 }catch(ignore){result.errors.push('อ่านรายการข้อความค้างไม่ได้ แต่ยังแสดงผลโควตาจาก LINE ได้');}
 return result;
}
// Run from Apps Script editor for diagnosis even before updating the web. Does not send LINE.
function lbCheckLineQuota(){var result=lbLineQuota_();Logger.log(JSON.stringify(result,null,2));return result;}
