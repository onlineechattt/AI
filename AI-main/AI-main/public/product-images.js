// Authenticated session memory only. No private photos in persistent browser storage.
export function createImageCache(ttl=600000, clock=()=>Date.now()) {
 const entries=new Map(); let generation=0;
 return {
  get(p){const e=entries.get(p.sku);return e&&e.version===p.imageVersion&&clock()-e.at<ttl?e.image:'';},
  async load(products,fetchImages){
   const token=generation, missing=products.filter(p=>!this.get(p));
   if(!missing.length)return;
   const rows=await fetchImages(missing.map(p=>p.sku));
   if(token!==generation)return;
   for(const row of rows||[])if(row.image)entries.set(row.sku,{image:row.image,version:row.imageVersion,at:clock()});
  },
  clear(){generation++;entries.clear();}
 };
}
