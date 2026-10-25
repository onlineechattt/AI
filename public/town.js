import {townMap} from './town-map.js';
import {townState} from './town-state.js';

const captions={cashier:'พร้อมรับออเดอร์ค่ะ',packer:'ดูแลชั้นสินค้า',customer:'แวะมาที่ร้านค่ะ'};
function actor(role){return `<div class="town-person ${role}" aria-hidden="true"><div class="person-shadow"></div><div class="person-sprite"></div><span class="person-bubble">${captions[role]}</span></div>`;}

export function createTown(root,{onAction,onPause=()=>{}}={}) {
  root.className='town';
  root.innerHTML=`<div class="town-heading"><div><span class="town-kicker">LILBULLE LITTLE TOWN</span><h2>เมืองเล็ก ๆ ที่ดูแลกัน</h2></div><div class="town-tools"><span class="town-open" id="townOpen">รอเปิดร้าน</span><button id="townPause" type="button" aria-pressed="false">พักการเคลื่อนไหว</button></div></div>
  <div class="town-viewport" tabindex="0" aria-label="แผนที่เมือง lilbulle เลื่อนดูซ้ายขวาได้"><div class="town-stage">
  <div class="town-cloud cloud-one"></div><div class="town-cloud cloud-two"></div>
  ${townMap}
  <div class="town-coordinates" aria-hidden="true"><b>lilbulle</b><span>little town · big care</span></div>
  <div class="town-shelf" id="townShelf" aria-label="รูปสินค้าจริงบนชั้น"><span>สินค้าจริงของเรา</span></div>
  ${actor('cashier')}${actor('packer')}${actor('customer')}
  <button class="town-pin pin-shop" data-town-action="shop"><span class="pin-number">01</span><span>ร้าน lilbulle<small>เลือกสินค้า / รับออเดอร์</small></span><span class="pin-arrow">↗</span></button>
  <button class="town-pin pin-warehouse" data-town-action="stock"><span class="pin-number">02</span><span>คลังและจัดของ<small>เติม / แจก / เทสเตอร์</small></span><span class="pin-arrow">↗</span></button>
  <button class="town-pin pin-market" data-town-action="booth"><span class="pin-number">03</span><span>งานบูธของเรา<small id="townBoothCount">เลือกงานที่กำลังขาย</small></span><span class="pin-arrow">↗</span></button>
  <button class="town-pin pin-park" data-town-action="history"><span class="pin-number">04</span><span>สมุดร้าน<small>ประวัติขาย / ตรวจบิลซ้ำ</small></span><span class="pin-arrow">↗</span></button>
  <div class="town-success" aria-hidden="true"><span>♡</span><b>ขอบคุณค่ะ</b><small>บันทึกขายสำเร็จ</small></div>
  <div class="town-view-note">เลื่อนดูเมือง · แตะป้ายเพื่อเปิดเมนู</div>
  </div></div>
  <div class="town-footer"><div class="town-story"><i></i><p id="townCaption" aria-live="polite">เลือกงานบูธ แล้วเปิดเมืองเล็กของเรา</p></div><div class="town-team"><span><i class="team-dot blue"></i>รับออเดอร์</span><span><i class="team-dot sand"></i>จัดสินค้า</span><span class="town-simulation">ฉากจำลองตามสถานะ POS</span></div></div>
  <div class="town-booths" id="townBooths" aria-label="เลือกงานบูธในเมือง"></div>`;
  let snapshot={},feedback='',feedbackTimer=null,paused=false,inView=true,lastShelf='',disposed=false;
  const pauseButton=root.querySelector('#townPause');
  const viewport=root.querySelector('.town-viewport');
  const media=window.matchMedia('(prefers-reduced-motion: reduce)');
  const updateMotion=()=>{root.dataset.paused=String(paused||media.matches||document.hidden||!inView);pauseButton.textContent=paused?'ให้เมืองเคลื่อนไหว':'พักการเคลื่อนไหว';pauseButton.setAttribute('aria-pressed',String(paused));};
  const visibility=()=>updateMotion();document.addEventListener('visibilitychange',visibility);
  media.addEventListener?.('change',updateMotion);
  const observer=typeof IntersectionObserver!=='undefined'?new IntersectionObserver(entries=>{inView=entries[0].isIntersecting;updateMotion();},{threshold:.05}):null;
  observer?.observe(root);
  pauseButton.addEventListener('click',()=>{paused=!paused;updateMotion();onPause(paused);});
  root.addEventListener('click',e=>{const button=e.target.closest('[data-town-action]');if(button&&!button.disabled)onAction?.(button.dataset.townAction,button.dataset.issue);});

  function render(){
    if(disposed)return;
    const derived=townState({...snapshot,feedback});root.dataset.mode=derived.mode;
    root.querySelector('#townCaption').textContent=derived.text;
    const state=snapshot.state;
    root.querySelector('#townOpen').textContent=!state?'รอเลือกบูธ':state.closed?'จบงานแล้ว':state.dayClosed?'ปิดยอดวันนี้':state.received?'เปิดร้านแล้ว':'รอตรวจรับ';
    root.querySelector('#townBoothCount').textContent=(snapshot.booths?.length||0)+' งาน · เลือกสลับบูธ';
    root.querySelector('.cashier .person-bubble').textContent=['order','packing'].includes(derived.mode)?'รับออเดอร์แล้วค่ะ':derived.mode==='paid'?'ขอบคุณค่ะ ♡':'พร้อมรับออเดอร์ค่ะ';
    root.querySelector('.packer .person-bubble').textContent=['packing','processing','pending'].includes(derived.mode)?'กำลังจัดของค่ะ':'ดูแลชั้นสินค้า';
    root.querySelector('.customer .person-bubble').textContent=derived.mode==='pending'?'รอยืนยันบิลค่ะ':'ขอสั่งสินค้าค่ะ';
    const locked=!!snapshot.busy||!!snapshot.pending;
    root.querySelectorAll('[data-town-action]').forEach(b=>{b.disabled=locked;});
    const shelf=root.querySelector('#townShelf');
    const products=(snapshot.products||[]).filter(p=>state?.lots.some(l=>l.sku===p.sku)).slice(0,5);
    const key=JSON.stringify(products.map(p=>[p.sku,p.image]));
    if(lastShelf!==key){lastShelf=key;shelf.replaceChildren();
      for(const p of products){const image=String(p.image||'');
        if(!/^data:image\/(jpeg|png|webp);base64,/.test(image)&&!/^https:\/\//.test(image))continue;
        if(/^https:\/\/(?:drive|docs)\.google\.com\//.test(image))continue;
        const img=document.createElement('img');img.src=image;img.alt=p.name+' · รูปจริง';img.decoding='async';img.referrerPolicy='no-referrer';img.addEventListener('error',()=>img.remove(),{once:true});shelf.append(img);
      }
      if(!shelf.children.length){const text=document.createElement('span');text.textContent='รูปสินค้าจริงจะแสดงที่นี่';shelf.append(text);}
    }
    const booths=root.querySelector('#townBooths');
    const boothKey=JSON.stringify((snapshot.booths||[]).map(b=>[b.id,b.name]))+'|'+(state?.id||'');
    if(booths.dataset.key!==boothKey){booths.dataset.key=boothKey;booths.replaceChildren();
      for(const [i,b] of (snapshot.booths||[]).entries()){
        const button=document.createElement('button');button.className='town-booth-card';button.type='button';button.dataset.townAction='selectBooth';button.dataset.issue=b.id;button.setAttribute('aria-pressed',String(state?.id===b.id));
        const number=document.createElement('span');number.className='booth-card-number';number.textContent=String(i+1).padStart(2,'0');
        const info=document.createElement('span');const name=document.createElement('b');name.textContent=b.name;const id=document.createElement('small');id.textContent=b.id;info.append(name,id);
        const dot=document.createElement('i');dot.className='booth-active-dot';button.append(number,info,dot);button.disabled=locked;booths.append(button);
      }
    }
    updateMotion();
  }
  requestAnimationFrame(()=>{viewport.scrollLeft=Math.max(0,(viewport.scrollWidth-viewport.clientWidth)/2);});
  render();
  return {
    update(value){snapshot=value;render();},
    event(kind){clearTimeout(feedbackTimer);feedback=['paid','restock'].includes(kind)?kind:'';render();if(feedback)feedbackTimer=setTimeout(()=>{feedback='';render();},4500);},
    clear(){clearTimeout(feedbackTimer);feedback='';snapshot={};lastShelf='';render();},
    destroy(){disposed=true;clearTimeout(feedbackTimer);observer?.disconnect();document.removeEventListener('visibilitychange',visibility);media.removeEventListener?.('change',updateMotion);}
  };
}
