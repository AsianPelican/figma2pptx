// Export a PPTX to PDF with PowerPoint for Mac, from a temp copy. Returns slide count.
import {execFileSync} from 'node:child_process';import {renameSync,copyFileSync,mkdirSync,writeFileSync,rmSync} from 'node:fs';import {join,resolve,basename} from 'node:path';import {tmpdir} from 'node:os';
export function ppExport(pptx:string,pdf:string):number{const tmp=join(tmpdir(),`f2p-${Date.now()}`);mkdirSync(tmp,{recursive:true});const deck=join(tmp,`F2P_${Date.now()}_${basename(pptx)}`);copyFileSync(resolve(pptx),deck);
 const script=`on run argv
 set s to POSIX file (item 1 of argv)
 set o to POSIX file (item 2 of argv)
 set nm to name of (info for s)
 tell application "Microsoft PowerPoint"
  open s
  repeat 300 times
   if exists presentation nm then exit repeat
   delay 0.2
  end repeat
  set d to presentation nm
  set c to count of slides of d
  with timeout of 1800 seconds
   save d in o as save as PDF
  end timeout
  close d saving no
  return c
 end tell
end run`;writeFileSync(join(tmp,'x.applescript'),script);
 const staged=resolve(`_export_${Date.now()}.pdf`);const n=+execFileSync("osascript",[join(tmp,"x.applescript"),deck,staged],{encoding:'utf8',timeout:600000}).trim();rmSync(tmp,{recursive:true,force:true});renameSync(staged,resolve(pdf));return n}
if(import.meta.main){const t=Date.now();console.log('slides',ppExport(process.argv[2],process.argv[3]),'secs',((Date.now()-t)/1000).toFixed(1))}
