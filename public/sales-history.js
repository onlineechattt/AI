export function lbHistoryRows_(events,booths,p){
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
