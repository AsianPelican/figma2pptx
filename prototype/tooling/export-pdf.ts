// Repeatable PPTX -> PDF export with size optimization and a size log.
//
//   bun export-pdf.ts <deck.pptx> [--preset screen|standard|print|all] [--out <dir>]   (default: screen, the captain's pick 2026-09-24)
//
// 1. Hashes the pptx, copies it to a uniquely named temp file (the source is never opened by PowerPoint,
//    so a locked deck stays byte-identical), exports it with Microsoft PowerPoint "save as PDF", and checks
//    the page count against the slide count.
// 2. Optimizes the PDF per preset, touching only images:
//    - photos (DCT/JPEG): downsampled to the preset ppi at their largest on-page size, re-encoded at the
//      preset JPEG quality; the smaller of old/new is kept. Only opaque plain JPEGs are touched: anything with
//      a soft/stencil mask or Decode array stays byte-identical (PowerPoint stores some headline rasters that way).
//    - lossless images (Flate, mostly PowerPoint's 300 ppi rasters of the outlined PP Formula headlines,
//      logos and diagram strokes): never downsampled and never turned into JPEG, so type stays crisp;
//      re-deflated at level 9 with per-row PNG predictors, which is bit-for-bit lossless.
//    - identical objects deduplicated, all streams Flate-compressed (mupdf garbage=4).
//    Text, vector paths and fonts are not rewritten; extracted text must match the raw export exactly.
// 3. Appends one line per output to size-log.jsonl next to the output:
//    date, source path and sha256, pages, raw MB, optimized MB, preset settings.
import * as m from 'mupdf';
import {readFileSync,writeFileSync,copyFileSync,rmSync,mkdirSync,appendFileSync,existsSync,statSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {resolve,basename,dirname,join} from 'node:path';
import {tmpdir} from 'node:os';
import {deflateSync} from 'node:zlib';

// Lossless re-deflate with per-row PNG predictors (None/Sub/Up/Average/Paeth, minimum-sum heuristic).
function pngFilter(px:Uint8Array,w:number,h:number,bpp:number){const stride=w*bpp,out=new Uint8Array((stride+1)*h);const prev=new Uint8Array(stride);
 for(let y=0;y<h;y++){const row=px.subarray(y*stride,(y+1)*stride);let best=0,bestSum=Infinity,bestBuf:Uint8Array|null=null;
  for(let t=0;t<5;t++){const b=new Uint8Array(stride);let sum=0;for(let x=0;x<stride;x++){const a=x>=bpp?row[x-bpp]:0,up=prev[x],c=x>=bpp?prev[x-bpp]:0;let v=row[x];
    if(t===1)v-=a;else if(t===2)v-=up;else if(t===3)v-=(a+up)>>1;else if(t===4){const p=a+up-c,pa=Math.abs(p-a),pb=Math.abs(p-up),pc=Math.abs(p-c);v-=pa<=pb&&pa<=pc?a:pb<=pc?up:c}
    b[x]=v&255;sum+=b[x]<128?b[x]:256-b[x]}if(sum<bestSum){bestSum=sum;best=t;bestBuf=b}}
  out[y*(stride+1)]=best;out.set(bestBuf!,y*(stride+1)+1);prev.set(row)}
 return out}

const PRESETS:Record<string,{photoPpi:number,jpegQ:number,chroma:string}>={
 screen:{photoPpi:150,jpegQ:80,chroma:'4:2:0'},
 standard:{photoPpi:200,jpegQ:85,chroma:'4:2:0'},
 print:{photoPpi:300,jpegQ:90,chroma:'4:4:4'},
};
const args=process.argv.slice(2);const src=resolve(args[0]||'');
if(!args[0]||!existsSync(src))throw Error('usage: bun export-pdf.ts <deck.pptx|deck.pdf> [--preset screen|standard|print|all] [--out <dir>]');
const maskedPhotos=args.includes('--masked-photos');
const fromPdf=/\.pdf$/i.test(src); // a finished PDF (e.g. merged Figma exports): skip PowerPoint, optimize it directly
const opt=(k:string,d:string)=>{const i=args.indexOf(k);return i>=0?args[i+1]:d};
const presetArg=opt('--preset','screen'),outDir=resolve(opt('--out',dirname(src)));mkdirSync(outDir,{recursive:true});
const presets=presetArg==='all'?Object.keys(PRESETS):[presetArg];
for(const p of presets)if(!PRESETS[p])throw Error('unknown preset '+p);
const sha=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex');
const MB=(n:number)=>+(n/1e6).toFixed(2);
const stem=basename(src).replace(/\.(pptx|pdf)$/i,'');
const srcBytes=readFileSync(src),srcHash=sha(srcBytes);

// 1. Export from a temp copy.
let slides=0;const rawPdf=join(outDir,`${stem}_raw-export.pdf`);
if(fromPdf){copyFileSync(src,rawPdf);slides=m.Document.openDocument(Buffer.from(readFileSync(rawPdf)),'application/pdf').countPages()}else{
const tmp=join(tmpdir(),`pdfexport-${Date.now()}`);mkdirSync(tmp,{recursive:true});
const tmpDeck=join(tmp,`Export_${Date.now()}_${stem}.pptx`);
copyFileSync(src,tmpDeck);
const script=`on run argv
 set s to POSIX file (item 1 of argv)
 set o to POSIX file (item 2 of argv)
 set nm to name of (info for s)
 tell application "Microsoft PowerPoint"
  open s
  repeat 150 times
   if exists presentation nm then exit repeat
   delay 0.2
  end repeat
  set d to presentation nm
  set c to count of slides of d
  save d in o as save as PDF
  close d saving no
  return c
 end tell
end run`;
writeFileSync(join(tmp,'export.applescript'),script);
slides=+execFileSync('osascript',[join(tmp,'export.applescript'),tmpDeck,rawPdf],{encoding:'utf8',timeout:300000}).trim();
rmSync(tmp,{recursive:true,force:true});
if(sha(readFileSync(src))!==srcHash)throw Error('source pptx changed during export');}
const rawBytes=readFileSync(rawPdf);
const pagesOf=(b:Uint8Array)=>m.Document.openDocument(Buffer.from(b),'application/pdf').countPages();
if(pagesOf(rawBytes)!==slides)throw Error(`page count ${pagesOf(rawBytes)} != slides ${slides}`);

// Largest on-page size of every image object, from the raw export.
const imgUse=new Map<number,{minPpi:number}>();
for(const line of execFileSync('pdfimages',['-list',rawPdf],{encoding:'utf8'}).trim().split('\n').slice(2)){
 const c=line.trim().split(/\s+/);const obj=+c[10],ppi=Math.min(+c[12],+c[13]);
 const u=imgUse.get(obj);if(!u||ppi<u.minPpi)imgUse.set(obj,{minPpi:ppi});
}
const textOf=(b:Uint8Array)=>{const d=m.Document.openDocument(Buffer.from(b),'application/pdf');let t='';for(let i=0;i<d.countPages();i++)t+=d.loadPage(i).toStructuredText().asText()+'\f';return t};
const rawText=textOf(rawBytes);

const results:any[]=[];const gateLog:any[]=[];
// Render gate: every page of raw and optimized is rendered at 150 dpi with poppler (independent of mupdf, which wrote
// the file) and compared in 12x12 px blocks (mean absolute difference, 0-255).
// Page areas (150 dpi px) covered by photos this run re-encoded, found by their new pixel size. Only there is
// JPEG/resampling difference expected; everywhere else the page must render essentially identically.
function photoZones(pdf:Uint8Array,dims:Set<string>){const d=m.Document.openDocument(Buffer.from(pdf),'application/pdf');const zones:Record<number,number[][]>={};const k=150/72;
 for(let p=0;p<d.countPages();p++){const page=d.loadPage(p);const list:number[][]=[];
  const dev=new m.Device({fillImage(img:any,ctm:number[]){if(!dims.has(`${img.getWidth()}x${img.getHeight()}`))return;const xs=[ctm[4],ctm[0]+ctm[4],ctm[2]+ctm[4],ctm[0]+ctm[2]+ctm[4]],ys=[ctm[5],ctm[1]+ctm[5],ctm[3]+ctm[5],ctm[1]+ctm[3]+ctm[5]];list.push([Math.min(...xs)*k,Math.min(...ys)*k,Math.max(...xs)*k,Math.max(...ys)*k])}} as any);
  page.run(dev,m.Matrix.identity);dev.close();zones[p+1]=list}
 return zones}
function renderGate(a:string,b:string,zones:Record<number,number[][]>){const dir=join(tmpdir(),`gate-${Date.now()}`);mkdirSync(dir,{recursive:true});
 execFileSync('pdftoppm',['-r','150',a,join(dir,'a')]);execFileSync('pdftoppm',['-r','150',b,join(dir,'b')]);
 const ppm=(f:string)=>{const d=readFileSync(f);let o=0,fields:number[]=[];while(fields.length<3){while(d[o]===35||d[o]<=32){if(d[o]===35)while(d[o]!==10)o++;o++}let s='';while(d[o]>32)s+=String.fromCharCode(d[o++]);if(s!=='P6')fields.push(+s)}o++;return{w:fields[0],h:fields[1],px:d.subarray(o)}};
 const pages:any[]=[];const failed:number[]=[];
 for(const f of require('node:fs').readdirSync(dir).filter((x:string)=>x.startsWith('a-')).sort()){const n=+f.match(/(\d+)\.ppm$/)![1];const A=ppm(join(dir,f)),B=ppm(join(dir,f.replace(/^a-/,'b-')));
  if(A.w!==B.w||A.h!==B.h){pages.push({page:n,pass:false,reason:'size'});failed.push(n);continue}
  const bs=12;let worst=0,over8=0,blocks=0,worstPhoto=0;const Z=zones[n]||[];const inZone=(x:number,y:number)=>Z.some(z=>x+bs>z[0]&&y+bs>z[1]&&x<z[2]&&y<z[3]);
  for(let by=0;by<A.h;by+=bs)for(let bx=0;bx<A.w;bx+=bs){let sum=0,cnt=0;for(let y=by;y<Math.min(by+bs,A.h);y++)for(let x=bx;x<Math.min(bx+bs,A.w);x++){const k=(y*A.w+x)*3;sum+=Math.abs(A.px[k]-B.px[k])+Math.abs(A.px[k+1]-B.px[k+1])+Math.abs(A.px[k+2]-B.px[k+2]);cnt+=3}
   const mean=sum/cnt;if(inZone(bx,by)){if(mean>worstPhoto)worstPhoto=mean;continue}blocks++;if(mean>worst)worst=mean;if(mean>8)over8++}
  // Outside re-encoded photos: must be identical to within anti-aliasing (worst block <= 6/255).
  // Inside them: resampling noise is fine, gross damage (a block off by > 90/255, e.g. a lost or black image) is not.
  const pass=worst<=6&&over8===0&&worstPhoto<=90;pages.push({page:n,worstBlockOutsidePhotos:+worst.toFixed(1),worstBlockInsideResampledPhotos:+worstPhoto.toFixed(1),photoZones:Z.length,pass});if(!pass)failed.push(n)}
 rmSync(dir,{recursive:true,force:true});return{pass:failed.length===0,failed,pages}}
for(const name of presets){
 const P=PRESETS[name];
 const doc=m.Document.openDocument(Buffer.from(rawBytes),'application/pdf') as m.PDFDocument;
 const done=new Set<number>();let photos=0,photosResized=0,lossless=0;const changedDims=new Set<string>();
 for(let i=1;i<doc.countObjects();i++){
  const ref=doc.newIndirect(i,0);if(!ref.isStream())continue;
  // Read through the stream ref only: resolve() on every object makes mupdf repair the Quartz xref and
  // silently drops image references on save.
  const d=ref;const st=d.get('Subtype');if(st.isNull()||st.asName()!=='Image')continue;
  const f=d.get('Filter');const filter=f.isNull()?'':f.isArray()?'multi':f.asName();
  if(filter==='FlateDecode'){ // lossless: same pixels, stronger deflate
   const w=d.get('Width').asNumber(),h=d.get('Height').asNumber(),bpc=d.get('BitsPerComponent').asNumber();
   const raw=ref.readRawStream().getLength(),px=ref.readStream().asUint8Array(),bpp=px.length/(w*h);
   if(bpc!==8||!Number.isInteger(bpp))continue;
   const plain=deflateSync(px,{level:9}),pred=deflateSync(pngFilter(px,w,h,bpp),{level:9});
   const usePred=pred.length<plain.length,best=usePred?pred:plain;if(best.length>=raw)continue;
   ref.writeRawStream(best);
   if(usePred){const dp=doc.newDictionary();dp.put('Predictor',15);dp.put('Colors',bpp);dp.put('BitsPerComponent',8);dp.put('Columns',w);ref.put('DecodeParms',dp)}else ref.delete('DecodeParms');
   lossless++;continue}
  if(filter!=='DCTDecode')continue;
  // Only opaque, plain photos: anything with a soft mask, stencil mask, Decode array or non-8-bit data is left
  // byte-for-byte alone (PowerPoint also stores some headline rasters as JPEG + soft mask).
  // Masked photos (Figma attaches alpha to every image) are allowed only with --masked-photos and only when large:
  // the JPEG is resampled/re-encoded and its soft mask left byte-identical (a PDF soft mask need not share the
  // image's pixel size). The render gate below still has the final say.
  const smasked=!d.get('SMask').isNull();
  if(smasked&&!(maskedPhotos&&d.get('Width').asNumber()*d.get('Height').asNumber()>=500000))continue;
  if(!d.get('Mask').isNull()||!d.get('Decode').isNull()||(!d.get('ImageMask').isNull()&&d.get('ImageMask').asBoolean()))continue;
  if(d.get('BitsPerComponent').asNumber()!==8)continue;
  const use=imgUse.get(i);if(!use)continue;
  const w=d.get('Width').asNumber(),h=d.get('Height').asNumber();
  const scale=Math.min(1,P.photoPpi/use.minPpi),nw=Math.max(1,Math.round(w*scale)),nh=Math.max(1,Math.round(h*scale));
  const bytes=ref.readRawStream().asUint8Array();
  const cs=d.get('ColorSpace');const cmyk=!cs.isNull()&&cs.toString().includes('CMYK');
  const out=execFileSync('magick',['jpeg:-',...(scale<1?['-filter','Lanczos','-resize',`${nw}x${nh}!`]:[]),'-strip','-sampling-factor',P.chroma,'-quality',String(P.jpegQ),...(cmyk?['-colorspace','CMYK']:[]),'jpeg:-'],{input:bytes,maxBuffer:80e6});
  photos++;
  if(out.length>=bytes.length&&scale===1)continue;
  const cspace=(b:Uint8Array)=>execFileSync('magick',['identify','-format','%[colorspace]','jpeg:-'],{input:b,encoding:'utf8'}).trim();
  if(cspace(out)!==cspace(bytes))continue; // colour model must match the PDF colour space
  ref.writeRawStream(out);ref.put('Width',nw);ref.put('Height',nh);if(scale<1)photosResized++;changedDims.add(`${nw}x${nh}`);
 }
 const buf=Buffer.from(doc.saveToBuffer('garbage=4,compress=yes').asUint8Array());
 const imgs=(p:string)=>execFileSync('pdfimages',['-list',p],{encoding:'utf8'}).trim().split('\n').length-2;
 const outPath=join(outDir,`${stem}_${name}.pdf`);writeFileSync(outPath,buf);
 if(imgs(outPath)!==imgs(rawPdf))throw Error(`image placements changed in ${name} output`);
 const zones=photoZones(buf,changedDims);
 const gate=renderGate(rawPdf,outPath,zones);gateLog.push({preset:name,...gate});writeFileSync(join(outDir,`${stem}_${name}_render-gate.json`),JSON.stringify(gateLog.at(-1),null,2));
 if(!gate.pass){if(process.env.KEEP_FAILED!=='1')rmSync(outPath);throw Error(`render gate FAILED for ${name}: pages ${gate.failed.join(', ')} differ beyond JPEG noise; output deleted`)}
 if(textOf(buf)!==rawText)throw Error(`text changed in ${name} output`);
 if(pagesOf(buf)!==slides)throw Error(`page count changed in ${name} output`);
 const rec={date:new Date().toISOString(),source:src,sourceSha256:srcHash,pages:slides,rawMB:MB(rawBytes.length),optimizedMB:MB(buf.length),preset:name,settings:{...P,jpegResample:'Lanczos',flate:'max, dedupe (mupdf garbage=4)'},photos,photosResized,losslessRecompressed:lossless,output:outPath};
 appendFileSync(join(outDir,'size-log.jsonl'),JSON.stringify(rec)+'\n');writeFileSync(join(outDir,`${stem}_${name}_render-gate.json`),JSON.stringify(gateLog.at(-1),null,2));results.push(rec);
 console.log(`${name.padEnd(9)} ${MB(rawBytes.length)} MB -> ${MB(buf.length)} MB  (${photosResized}/${photos} photos resampled, q${P.jpegQ}, ${P.photoPpi} ppi)  ${outPath}`);
}
