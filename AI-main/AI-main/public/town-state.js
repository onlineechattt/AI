// The town only describes POS state. It never records a sale or moves stock.
export function townState({state, pending, busy, cartCount=0, feedback=''}) {
  if (pending) return {mode:'pending', text:'กำลังตรวจสอบผลรายการ · ลูกค้ารอที่เคาน์เตอร์'};
  if (feedback==='paid') return {mode:'paid', text:'บันทึกขายสำเร็จแล้ว · ขอบคุณที่แวะมาค่ะ'};
  if (feedback==='restock') return {mode:'restock', text:'บันทึกรายการสินค้าแล้ว · ทีมงานกำลังจัดของ'};
  if (busy) return {mode:'processing', text:'กำลังทำรายการ · รอระบบยืนยันก่อนนะคะ'};
  if (!state) return {mode:'welcome', text:'เลือกงานบูธ แล้วเปิดเมืองเล็กของเรา'};
  if (state.closed || state.dayClosed) return {mode:'closed', text:state.closed?'งานบูธนี้จบแล้ว · ดูประวัติย้อนหลังได้':'ปิดยอดวันนี้แล้ว · พักร้านกันก่อนค่ะ'};
  if (!state.received) return {mode:'receiving', text:'ของมาถึงแล้ว · ตรวจรับให้ครบก่อนเปิดขาย'};
  if (cartCount>0) return {mode:cartCount>1?'packing':'order', text:cartCount>1?'กำลังจัดสินค้าตามออเดอร์ · รอยืนยันรับเงิน':'ลูกค้ามาแล้วค่ะ · เลือกสินค้าได้เลย'};
  return {mode:'ready', text:'ร้านเปิดแล้ว · ทีมงานกำลังดูแลเมืองของเรา'};
}
