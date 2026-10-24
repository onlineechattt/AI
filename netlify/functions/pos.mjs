import {createHmac, timingSafeEqual} from 'node:crypto';
const json=(data,status=200,headers={})=>new Response(JSON.stringify(data),{status,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store',...headers}});
const sign=(s,key)=>createHmac('sha256',key).update(s).digest('base64url');
export function readSession(cookie,key){try{const raw=(cookie||'').split(';').map(x=>x.trim()).find(x=>x.startsWith('lb_session='))?.slice(11);if(!raw)return null;const [data,signature]=raw.split('.');const expected=sign(data,key);if(!signature||signature.length!==expected.length||!timingSafeEqual(Buffer.from(signature),Buffer.from(expected)))return null;const session=JSON.parse(Buffer.from(data,'base64url'));return session.exp>Date.now()&&['boss','team'].includes(session.role)?session:null;}catch{return null;}}
export default async function handler(request,context={}){
 if(request.method!=='POST')return json({error:'ใช้ POST เท่านั้น'},405);
 const origin=request.headers.get('origin');if(origin&&origin!==new URL(request.url).origin)return json({error:'ไม่อนุญาตต้นทางนี้'},403);
 const env=name=>globalThis.Netlify?.env?.get(name)||process.env[name];
 const APPS_SCRIPT_URL=env('APPS_SCRIPT_URL'),POS_API_KEY=env('POS_API_KEY');if(!APPS_SCRIPT_URL||!POS_API_KEY)return json({error:'ยังไม่ได้เชื่อมระบบจริง กรุณาตั้งค่า APPS_SCRIPT_URL และ POS_API_KEY บน Netlify'},503);
 if(!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(APPS_SCRIPT_URL))return json({error:'URL Apps Script ไม่ถูกต้อง'},503);
 const key=env('POS_SESSION_SECRET')||POS_API_KEY;
 try{
  if(Number(request.headers.get('content-length'))>3100000)return json({error:'ไฟล์ใหญ่เกินไป'},413);
  const text=await request.text();if(text.length>3100000)return json({error:'ไฟล์ใหญ่เกินไป'},413);
  const body=JSON.parse(text),action=body.action,payload=body.payload||{};
  if(action==='logout')return json({success:true},200,{'Set-Cookie':'lb_session=; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=0'});
  if(!['login','init','receive','sale','void','closeDay','upload','catalogSave','retryNotify','restock','receiveRestock','stockOut'].includes(action))return json({error:'คำสั่งไม่รองรับ'},400);
  const session=readSession(request.headers.get('cookie'),key);
  if(action!=='login'&&!session)return json({error:'กรุณาเข้าสู่ระบบ'},401);
  if(action!=='login')payload.staff=session.staff;
  const response=await fetch(APPS_SCRIPT_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({posApi:'lilbulle-v11',apiKey:POS_API_KEY,action,payload,role:session?.role,clientKey:sign(context.ip||request.headers.get('x-nf-client-connection-ip')||'unknown',key)}),signal:AbortSignal.timeout(55000),redirect:'follow'});
  const result=await response.json();if(!result.success)return json({error:result.error||'ระบบปฏิเสธรายการ',rejected:result.rejected===true},400);
  if(action==='login'){
   const staff=String(payload.staff||'').trim().slice(0,80);if(!staff)return json({error:'ระบุชื่อพนักงาน'},400);
   const data=Buffer.from(JSON.stringify({role:result.role,staff,exp:Date.now()+8*3600000})).toString('base64url');
   return json({success:true,role:result.role,staff},200,{'Set-Cookie':`lb_session=${data}.${sign(data,key)}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=28800`});
  }
  return json({success:true,result:result.result,role:session.role,staff:session.staff});
 }catch{return json({error:'ยังยืนยันผลไม่ได้ กรุณาส่งคำขอเดิมซ้ำเพื่อเช็กผล ห้ามสร้างบิลใหม่แทน'},502);}
}
