export const demoProducts=[{sku:'balm',name:'Baby Rich Balm',price:290,image:''},{sku:'lotion',name:'Organic Lotion',price:390,image:''},{sku:'foam',name:'Baby Foam',price:350,image:''},{sku:'mozzie',name:'Mozzie Guard',price:190,image:''},{sku:'spray',name:'Sanitizer Spray',price:150,image:''}];
export const day=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Asia/Bangkok'});
export function seed(){return {products:structuredClone(demoProducts),booths:[{id:'DEMO-BOOTH',name:'ร้านตัวอย่าง · Baby Fair',location:'ข้อมูลสมมติสำหรับทดลอง'}],received:false,events:[],lots:demoProducts.map(p=>({key:p.sku+'|31/12/2028',sku:p.sku,exp:'31/12/2028',issued:24,returned:0,sold:0,remaining:24}))};}
export function demoRun(db,action,p={}){
 if(action==='init')return snapshot(db,p.issueNo);
 if(action==='catalogSave'){const product=db.products.find(x=>x.sku===p.sku);Object.assign(product,{price:Number(p.price),image:p.image});return {success:true};}
 if(action==='retryNotify')return {success:true};
 if(action==='upload')return {success:true,id:'DEMO-PHOTO',data:{image:p.kind==='product'?`data:${p.mime};base64,${p.base64}`:'',url:''}};
 const old=db.events.find(e=>e.requestId===p.requestId);if(old)return {success:true,id:old.id,duplicate:true};
 if(p.issueNo!=='DEMO-BOOTH')throw Error('เลือกบูธตัวอย่างก่อน');
 const state=snapshot(db,p.issueNo).state;
 if(state.dayClosed)throw Error('ปิดยอดวันนี้แล้ว');
 let data,type;
 if(action==='receive'){if(db.received)throw Error('ตรวจรับแล้ว');if(p.counts.length!==db.lots.length||p.counts.some(x=>Number(x.actual)!==db.lots.find(l=>l.key===x.key)?.issued))throw Error('ยอดนับไม่ตรงใบเบิก');db.received=true;type='RECEIVE';data={counts:p.counts};}
 else if(!db.received)throw Error('ตรวจรับก่อนขาย');
 else if(action==='sale'){
 const stock=structuredClone(state.lots),lines=[];let total=0;
 for(const x of p.items){if(!Number.isInteger(x.qty)||x.qty<=0||!Number.isFinite(x.price)||x.price<=0)throw Error('จำนวนหรือราคาไม่ถูกต้อง');const lot=stock.find(l=>l.sku===x.sku);if(!lot||x.qty>lot.remaining)throw Error('สินค้าไม่พอ');lot.remaining-=x.qty;lines.push({key:lot.key,sku:x.sku,name:db.products.find(y=>y.sku===x.sku).name,qty:x.qty,exp:lot.exp,price:x.price});total+=x.qty*x.price;}
 data={lines,payment:p.payment,total:Math.round(total*100)/100,note:p.note,photoId:p.photoId};type='SALE';
 }else if(action==='void'){const sale=state.sales.find(e=>e.id===p.billId);if(!sale||sale.day!==day())throw Error('ยกเลิกไม่ได้');data={billId:p.billId,note:p.note};type='VOID';}
 else if(action==='closeDay'){if(p.counts.length!==state.lots.length)throw Error('นับให้ครบ');const summary=snapshot(db,p.issueNo).summary;const counts=p.counts.map(c=>({...c,expected:state.lots.find(l=>l.key===c.key).remaining}));delete summary.sales;data={summary,counts,actualMoney:p.money};type='CLOSE';}
 else throw Error('คำสั่งไม่รองรับ');
 const id=type+'-'+crypto.randomUUID();db.events.push({id,requestId:p.requestId,type,issue:p.issueNo,day:day(),at:new Date().toISOString(),staff:p.staff,data,notify:'NONE'});return {success:true,id,data};
}
export function snapshot(db,issue){if(!issue)return {products:db.products,booths:db.booths,state:null,summary:null,pendingNotifications:0};const voids=new Set(db.events.filter(e=>e.type==='VOID').map(e=>e.data.billId));const sales=db.events.filter(e=>e.type==='SALE'&&!voids.has(e.id));const lots=db.lots.map(l=>({...l,sold:0}));for(const e of sales)for(const x of e.data.lines)lots.find(l=>l.key===x.key).sold+=x.qty;for(const l of lots)l.remaining=l.issued-l.returned-l.sold;const today=sales.filter(e=>e.day===day()),money={CASH:0,TRANSFER:0,CARD:0,OTHER:0},products={};for(const e of today){money[e.data.payment]+=e.data.total;for(const l of e.data.lines){products[l.sku]??={sku:l.sku,name:l.name,qty:0,total:0};products[l.sku].qty+=l.qty;products[l.sku].total+=l.qty*l.price;}}return {products:db.products,booths:db.booths,state:{id:issue,name:db.booths[0].name,day:day(),closed:false,received:db.received,lots,sales,closes:db.events.filter(e=>e.type==='CLOSE').map(e=>({day:e.day,staff:e.staff,data:e.data})),dayClosed:db.events.some(e=>e.type==='CLOSE'&&e.day===day())},summary:{day:day(),money,total:today.reduce((s,e)=>s+e.data.total,0),bills:today.length,products:Object.values(products),sales:today},pendingNotifications:0};}
