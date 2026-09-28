// Fetch per-frame SVG (live text, node ids), strip embedded base64 images to keep cache small.
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
const env=Object.fromEntries(readFileSync(process.env.FIGMA_ENV_FILE!,'utf8').split('\n').filter(l=>l.includes('=')).map(l=>[l.slice(0,l.indexOf('=')),l.slice(l.indexOf('=')+1).trim()]));
export const TOKEN=env.FIGMA_TOKEN;
export async function figma(path:string){for(let i=0;i<5;i++){const r=await fetch('https://api.figma.com/v1/'+path,{headers:{'X-Figma-Token':TOKEN}});if(r.status===429){await Bun.sleep(5000*(i+1));continue}if(!r.ok)throw Error(path+' '+r.status+' '+await r.text());return r.json()}throw Error('rate limited')}
if(import.meta.main){const key=process.argv[2];const ids=process.argv[3].split(',');
 const todo=ids.filter(id=>!existsSync(`cache/svg/${id.replace(':','-')}.svg`));
 for(let i=0;i<todo.length;i+=4){const chunk=todo.slice(i,i+4);const j:any=await figma(`images/${key}?ids=${chunk.join(',')}&format=svg&svg_outline_text=false&svg_include_node_id=true&svg_simplify_stroke=false`);
  await Promise.all(chunk.map(async id=>{const s=await (await fetch(j.images[id])).text();writeFileSync(`cache/svg/${id.replace(':','-')}.svg`,s.replace(/xlink:href="data:[^"]{200,}"/g,'xlink:href="data:stripped"').replace(/href="data:[^"]{200,}"/g,'href="data:stripped"'));console.log(id,s.length)}))}}
