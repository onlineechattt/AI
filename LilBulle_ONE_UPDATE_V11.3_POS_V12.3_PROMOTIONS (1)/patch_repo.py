"""Apply lil bulle promotional UI changes to an existing checkout of onlineechattt/AI (main at 2026-11-10)."""
from pathlib import Path
import sys
root=Path(sys.argv[1] if len(sys.argv)>1 else '.')
app=root/'public/app.js'; proxy=root/'netlify/functions/pos.mjs'
if not app.is_file() or not proxy.is_file():raise SystemExit('ไม่พบ public/app.js และ netlify/functions/pos.mjs ในโฟลเดอร์โปรเจกต์')
s=app.read_text(encoding='utf8')
old="function applyBundlePrices(rows){for(const b of bundles){const row=rows?.find(r=>r.id===b.id),fallback=bundleDefaults.find(x=>x.id===b.id);b.price=row?row.price:fallback.price;}}"
new="function applyBundlePrices(rows){if(Array.isArray(data.promotions)){bundles.splice(0,bundles.length,...data.promotions.filter(p=>p.active).map(p=>({id:p.id,name:p.name,price:p.price,components:p.components})));return;}for(const b of bundles){const row=rows?.find(r=>r.id===b.id),fallback=bundleDefaults.find(x=>x.id===b.id);b.price=row?row.price:fallback.price;}}"
if s.count(old)!=1:raise SystemExit('โครงสร้าง applyBundlePrices ไม่ตรงเวอร์ชันที่เตรียม หยุดอัปเดตเพื่อความปลอดภัย')
s=s.replace(old,new)
start=s.find('function bundlePriceDialog(){');end=s.find('const manualLineKey=',start)
if start<0 or end<start:raise SystemExit('ไม่พบจุดแทนหน้าแก้ราคาโปรโมชัน')
code=r'''
function bundlePriceDialog(){
 if(role!=='boss'){toast('เมนูนี้สำหรับบอสเท่านั้น');return;}
 const promos=data.promotions||[];
 const opts='<option value="">— สร้างโปรโมชั่นใหม่ —</option>'+promos.map(p=>`<option value="${esc(p.id)}">${esc(p.name)}${p.active?'':' (ปิดขาย)'}</option>`).join('');
 const productOpts='<option value="">— ไม่ใช้รายการนี้ —</option>'+data.products.map(p=>`<option value="${esc(p.sku)}">${esc(p.name)}</option>`).join('');
 showDialog('จัดการโปรโมชั่น',`<p class="hint">เฉพาะบอสสร้าง/เปลี่ยนชื่อ/จัดชุด/ราคา/เปิด-ปิด · ไม่แก้ยอด Stock · บิลเดิมคงราคาเดิม</p>
  <label>เลือกโปรโมชั่น<select id="promoSelect">${opts}</select></label>
  <label>ชื่อโปรโมชั่น<input id="promoName" maxlength="100" placeholder="เช่น คู่หูผิวนุ่ม"></label>
  <label>ราคาทั้งชุด (บาท)<input id="promoPrice" type="number" min="0.01" step="0.01" required></label>
  <div id="promoLines">${Array.from({length:5},(_,i)=>`<div class="catalog-edit"><label>สินค้า ${i+1}<select data-promo-sku="${i}">${productOpts}</select></label><label>จำนวน<input data-promo-qty="${i}" type="number" min="1" max="999" step="1" value="1"></label></div>`).join('')}</div>
  <label><input id="promoActive" type="checkbox" checked> เปิดขายโปรโมชั่นนี้</label>
  <button id="promoSave" class="primary">บันทึกโปรโมชั่น</button><p id="promoHint" class="hint">สินค้าชนิดเดียวกันให้รวมจำนวนในแถวเดียวกัน · เปลี่ยนราคาแล้วไม่แก้บิลเก่า</p>`);
 const select=$('promoSelect');
 function fill(){const p=promos.find(x=>x.id===select.value);$('promoName').value=p?.name||'';$('promoPrice').value=p?.price??'';$('promoActive').checked=p?p.active:true;for(let i=0;i<5;i++){document.querySelector(`[data-promo-sku="${i}"]`).value=p?.components?.[i]?.sku||'';document.querySelector(`[data-promo-qty="${i}"]`).value=p?.components?.[i]?.qty||1;}}
 select.onchange=fill;fill();
 $('promoSave').onclick=safe(async()=>{
  if(demo)throw Error('ต้องเข้าสู่ระบบจริงก่อนบันทึกโปรโมชั่น');
  const original=promos.find(x=>x.id===select.value),name=$('promoName').value.trim(),price=Number($('promoPrice').value),components=[];
  for(let i=0;i<5;i++){const sku=document.querySelector(`[data-promo-sku="${i}"]`).value;if(!sku)continue;components.push({sku,qty:Number(document.querySelector(`[data-promo-qty="${i}"]`).value)});}
  if(!name||!components.length||!(price>0))throw Error('กรอกชื่อ ราคา และสินค้าอย่างน้อยหนึ่งรายการ');
  const btn=$('promoSave');btn.disabled=true;try{
   const saved=await api('promotionSave',{id:original?.id||'',revision:original?.revision||'',name,price,components,active:$('promoActive').checked});
   data.promotions=saved.promotions;applyBundlePrices(data.bundlePrices);renderBundleButtons();toast('บันทึกโปรโมชั่นสำเร็จ · เครื่องอื่นรีเฟรชเพื่อรับข้อมูลใหม่');
   $('actionDialog').close();bundlePriceDialog();$('promoSelect').value=saved.id;$('promoSelect').dispatchEvent(new Event('change'));
  }finally{if(btn.isConnected)btn.disabled=false;}
 });
}

'''
s=s[:start]+code+s[end:]
# Keep admin button private, including cases where the role changes after login.
anchor="if($('bundlePricesBtn'))$('bundlePricesBtn').onclick=bundlePriceDialog;"
if s.count(anchor)!=1:raise SystemExit('ไม่พบปุ่มตั้งโปรโมชั่น')
s=s.replace(anchor,anchor+"if($('bundlePricesBtn'))$('bundlePricesBtn').hidden=role!=='boss';",1)
# In the UI draw step, also enforce role on every refresh.
render_anchor="function render(){const selected=issue;"
if s.count(render_anchor)!=1:raise SystemExit('โครงสร้าง render เปลี่ยนไป')
s=s.replace(render_anchor,"function render(){if($('bundlePricesBtn'))$('bundlePricesBtn').hidden=role!=='boss';const selected=issue;",1)
# Update version label only, no other layout changes.
s=s.replace("const APP_VERSION='12.1.15';","const APP_VERSION='12.3.0';",1)
proxy_src=proxy.read_text(encoding='utf8')
needle="'bundlePriceSave','retryNotify'"
if proxy_src.count(needle)!=1:raise SystemExit('ไม่พบ allowlist API ใน Netlify')
proxy_src=proxy_src.replace(needle,"'bundlePriceSave','promotionSave','retryNotify'",1)
app.write_text(s,encoding='utf8');proxy.write_text(proxy_src,encoding='utf8')
print('PATCH_OK public/app.js + netlify/functions/pos.mjs')
