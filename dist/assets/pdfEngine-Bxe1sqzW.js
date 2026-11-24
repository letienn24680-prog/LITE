import{g as y,T as L,R as A,p as N,G as S}from"./vendor-pdfjs-BC_cRcEG.js";S.workerSrc=N;const w=16777216,x="data-lite-pdf-text-layer",d=new WeakMap,m=new WeakMap;function C(e){return e instanceof A?!0:e instanceof Error?e.name==="RenderingCancelledException"||e.name==="AbortException"||e.message.toLowerCase().includes("rendering cancelled"):!1}function D(e){if(!Number.isFinite(e)||e<=0)return"I";const a=[[1e3,"M"],[900,"CM"],[500,"D"],[400,"CD"],[100,"C"],[90,"XC"],[50,"L"],[40,"XL"],[10,"X"],[9,"IX"],[5,"V"],[4,"IV"],[1,"I"]];let t=Math.floor(e),n="";for(const[r,s]of a)for(;t>=r;)n+=s,t-=r;return n}function E(e,a=260){const t=e.replace(/\s+/g," ").trim();return t.length<=a?t:`${t.slice(0,a-1).trimEnd()}…`}function z(){if(typeof document>"u"||document.head.querySelector(`style[${x}="true"]`))return;const e=document.createElement("style");e.setAttribute(x,"true"),e.textContent=`
.textLayer {
  position: absolute;
  text-align: initial;
  inset: 0;
  overflow: clip;
  opacity: 1;
  line-height: 1;
  letter-spacing: normal;
  word-spacing: normal;
  -webkit-text-size-adjust: none;
  text-size-adjust: none;
  forced-color-adjust: none;
  transform-origin: 0 0;
  caret-color: CanvasText;
  z-index: 2;
  --min-font-size: 1;
  --text-scale-factor: calc(var(--total-scale-factor, 1) * var(--min-font-size));
  --min-font-size-inv: calc(1 / var(--min-font-size));
}
.textLayer span,
.textLayer br {
  color: transparent;
  position: absolute;
  white-space: pre;
  cursor: text;
  transform-origin: 0% 0%;
  -webkit-user-select: text;
  user-select: text;
}
.textLayer > :not(.markedContent),
.textLayer .markedContent span:not(.markedContent) {
  z-index: 1;
  --font-height: 0;
  font-size: calc(var(--text-scale-factor) * var(--font-height));
  --scale-x: 1;
  --rotate: 0deg;
  transform: rotate(var(--rotate)) scaleX(var(--scale-x)) scale(var(--min-font-size-inv));
}
.textLayer .markedContent {
  display: contents;
}
.textLayer ::selection {
  background: rgba(79, 70, 229, 0.28);
  color: transparent;
}
`,document.head.appendChild(e)}function b(e,a){const t=e;return typeof t.destroy!="function"&&Object.defineProperty(t,"destroy",{value:()=>a(),configurable:!0,writable:!0}),t}async function $(e){if(typeof e=="string"){const r=y({url:e,withCredentials:!1}),s=await r.promise;return b(s,()=>r.destroy())}const a=e instanceof ArrayBuffer?new Uint8Array(e.slice(0)):e.slice(),t=y({data:a}),n=await t.promise;return b(n,()=>t.destroy())}async function P(e,a){if(!a)return 1;try{const t=typeof a=="string"?await e.getDestination(a):a;if(!Array.isArray(t)||t.length===0)return 1;const n=t[0];if(typeof n=="number"&&Number.isFinite(n))return Math.min(e.numPages,Math.max(1,Math.floor(n)+1));if(n&&typeof n=="object"&&"num"in n&&"gen"in n){const r=await e.getPageIndex(n);return Math.min(e.numPages,Math.max(1,r+1))}}catch{return 1}return 1}async function R(e){const a=await e.getOutline().catch(()=>null);if(!Array.isArray(a)||a.length===0)return[];const t=[],n=60;let r=0;for(const s of a){if(t.length>=n)break;const h=(s.title??"").replace(/\s+/g," ").trim();if(!h)continue;r++;const c=await P(e,s.dest),o=Array.isArray(s.items)?s.items.map(i=>(i.title??"").replace(/\s+/g," ").trim()).filter(Boolean):[];if(t.push({id:`pdf-outline-${t.length+1}`,romanIndex:D(r),title:h,summary:o.length>0?o.slice(0,2).join(" · "):`Trang ${c}`,keyPoints:o.slice(0,3),pageNumber:c}),Array.isArray(s.items)&&s.items.length>0){let i=0;for(const f of s.items){if(t.length>=n)break;const l=(f.title??"").replace(/\s+/g," ").trim();if(!l)continue;i++;const u=await P(e,f.dest);t.push({id:`pdf-outline-${t.length+1}`,romanIndex:`${r}.${i}`,title:l,summary:`Trang ${u}`,keyPoints:[],pageNumber:u})}}}return t}async function I(e){const a=await e.arrayBuffer(),t=await $(a);try{const n=t.numPages,r=[],s=Math.min(n,2);for(let o=1;o<=s;o++){const i=await t.getPage(o);try{const f=await i.getTextContent(),l=[];for(const u of f.items)"str"in u&&typeof u.str=="string"&&u.str.trim().length>0&&l.push(u.str);l.length>0&&r.push(l.join(" "))}finally{i.cleanup()}if(r.join(" ").length>=300)break}const h=await R(t),c=E(r.join(" "),260);return{pageCount:n,contentPreview:c,outline:h}}finally{await t.destroy().catch(()=>{})}}async function V(e){const{pdfDoc:a,canvas:t}=e,n=d.get(t);if(n){try{n.cancel()}catch{}d.delete(t)}const r=Math.min(Math.max(1,Math.floor(e.pageNumber)),a.numPages),s=await a.getPage(r),h=s.getViewport({scale:1}),c=typeof e.scale=="number"&&Number.isFinite(e.scale)&&e.scale>0?e.scale:1,o=typeof e.targetWidthCss=="number"&&Number.isFinite(e.targetWidthCss)&&e.targetWidthCss>0?e.targetWidthCss/Math.max(1,h.width)*c:c,i=s.getViewport({scale:o}),f=typeof window<"u"&&window.devicePixelRatio||1;let l=Math.min(f,e.maxDpr??2);const u=i.width*l*(i.height*l);u>w&&(l*=Math.sqrt(w/u));const M=Math.max(1,Math.floor(i.width*l)),v=Math.max(1,Math.floor(i.height*l));t.width=M,t.height=v,t.style.width=`${i.width}px`,t.style.height=`${i.height}px`;const T=l!==1?[l,0,0,l,0,0]:void 0,g=s.render({canvas:t,viewport:i,transform:T});d.set(t,g);const k=()=>{try{g.cancel()}catch{}d.get(t)===g&&d.delete(t)};try{await g.promise}catch(p){if(!C(p))throw p}finally{d.get(t)===g&&d.delete(t)}return{widthCss:i.width,heightCss:i.height,cancel:k}}async function W(e){const{pdfDoc:a,container:t}=e;z();const n=m.get(t);if(n){try{n.cancel()}catch{}m.delete(t)}t.replaceChildren();const r=Math.min(Math.max(1,Math.floor(e.pageNumber)),a.numPages),s=Number.isFinite(e.viewportScale)&&e.viewportScale>0?e.viewportScale:1,h=await a.getPage(r),c=h.getViewport({scale:s});t.classList.add("textLayer"),t.style.setProperty("--scale-factor",String(c.scale)),t.style.setProperty("--total-scale-factor",String(c.scale)),t.style.setProperty("--scale-round-x","1px"),t.style.setProperty("--scale-round-y","1px"),t.style.width=`${c.width}px`,t.style.height=`${c.height}px`;const o=await h.getTextContent(),i=new L({textContentSource:o,container:t,viewport:c});m.set(t,i);let f=!1;const l=()=>{if(!f){f=!0;try{i.cancel()}catch{}m.get(t)===i&&m.delete(t),t.replaceChildren()}};try{await i.render()}catch(u){if(!C(u))throw u}return l}async function X(e,a,t=160){const n=Math.min(Math.max(1,Math.floor(a)),e.numPages),r=await e.getPage(n),s=r.getViewport({scale:1}),h=Math.max(.1,t/Math.max(1,s.width)),c=r.getViewport({scale:h}),o=document.createElement("canvas");o.width=Math.max(1,Math.round(c.width)),o.height=Math.max(1,Math.round(c.height));try{await r.render({canvas:o,viewport:c}).promise;const f=o.toDataURL("image/webp",.82);return f.startsWith("data:image/webp")?f:o.toDataURL("image/png")}finally{o.width=0,o.height=0,r.cleanup()}}export{W as a,X as b,I as i,$ as l,V as r};
