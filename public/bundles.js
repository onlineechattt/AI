export const bundles=[
 {id:'balm-lotion',name:'บาม + โลชั่น',price:690,components:[{sku:'balm',qty:1},{sku:'lotion',qty:1}]},
 {id:'balm-duo',name:'บาม 2 ชิ้น',price:550,components:[{sku:'balm',qty:2}]},
 {id:'mozzie-duo',name:'สเปรย์กันยุง 2 ชิ้น',price:350,components:[{sku:'mozzie',qty:2}]},
 {id:'foam-duo',name:'โฟม 2 ชิ้น',price 490,components:[{sku:'foam',qty:2}]},
 {id:'spray-duo',name:'แอลกอฮอล์ 2 ชิ้น',price150,components:[{sku:'spray',qty:2}]},
 {id:'set-1',name:'ชุด SET 1 · ครบ 5 ตัว',price 1290,components:[{sku:'balm',qty:1},{sku:'lotion',qty:1},{sku:'mozzie',qty:1},{sku:'foam',qty:1},{sku:'spray',qty:1}]}
];
export function cents(value){if(value===''||value===null||!Number.isFinite(Number(value))||Number(value)<=0||Math.abs(Number(value)*100-Math.round(Number(value)*100))>0.000001)throw Error('กรอกราคามากกว่า 0 และทศนิยมไม่เกิน 2 ตำแหน่ง');const n=Math.round(Number(value)*100);if(!Number.isSafeInteger(n)||n>100000000)throw Error('ราคาเกินกำหนด');return n;}
export function cartDemand(cart){const result={};for(const row of cart){if(!Number.isInteger(row.qty)||row.qty<1||row.qty>100000)throw Error('จำนวนไม่ถูกต้อง');for(const c of row.components||[{sku:row.sku,qty:1}])result[c.sku]=(result[c.sku]||0)+c.qty*row.qty;}return result;}
export function assertStock(cart,available){for(const [sku,qty] of Object.entries(cartDemand(cart)))if(qty>available(sku))throw Error('สินค้าในบูธไม่พอสำหรับรายการนี้: '+sku);}
export function expandCart(cart){const grouped=new Map();const add=(sku,qty,priceCents)=>{const key=sku+':'+priceCents,old=grouped.get(key);if(old)old.qty+=qty;else grouped.set(key,{sku,qty,price:priceCents/100});};for(const row of cart){const amount=cents(row.price);if(!Number.isInteger(row.qty)||row.qty<1||row.qty>100000)throw Error('จำนวนไม่ถูกต้อง');if(!row.components){add(row.sku,row.qty,amount);continue;}const units=row.components.flatMap(c=>Array(c.qty).fill(c.sku));if(amount<units.length)throw Error('ราคาเซตต่ำเกินไปสำหรับจำนวนสินค้า');const base=Math.floor(amount/units.length),remainder=amount%units.length;units.forEach((sku,i)=>add(sku,row.qty,base+(i<remainder?1:0)));}const rows=[...grouped.values()];if(rows.length>30)throw Error('รายการเกินกำหนด กรุณาแยกบิล');return rows;}
export function cartTotal(cart){return Math.round(cart.reduce((sum,r)=>sum+Math.round((Number(r.price)||0)*100)*r.qty,0))/100;}
export function saleNote(cart,note=''){const sets=cart.filter(r=>r.components);const text=(sets.length?'เซต: '+sets.map(r=>r.name+' ×'+r.qty+' @'+Number(r.price).toFixed(2)).join(' / ')+(note?' | ':''):'')+note;if(text.length>500)throw Error('หมายเหตุพร้อมรายการเซตยาวเกิน 500 ตัวอักษร กรุณาย่อหมายเหตุหรือแยกบิล');return text;}
