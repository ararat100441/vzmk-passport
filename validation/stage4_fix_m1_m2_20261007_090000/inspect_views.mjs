// Read-only visual inspection of the changed profile geometry/status/notes.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import net from 'node:net';
import http from 'node:http';
import https from 'node:https';
let attempts=0;
function denied(){attempts++;throw new Error('Offline visual inspection');}
net.connect=net.createConnection=net.Socket.prototype.connect=denied;
http.get=http.request=https.get=https.request=globalThis.fetch=denied;
const out=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(out,'../..');
const require=createRequire(path.join(root,'profiles_stage4/workbook.mjs'));
const {FileBlob,SpreadsheetFile}=await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const workbook=path.join(out,'rendered_template.xlsx');
const data=JSON.parse(await fs.readFile(path.join(out,'baseline_generation/profiles.catalog.json'),'utf8'));
const identity=[...new Set(data.records.map(r=>r.key+'\0'+r.standard))].sort();
const row=identity.indexOf('L100×8\0ГОСТ 8509-93')+2;
if(row<2)throw new Error('Deterministic changed row missing');
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(workbook));
for(const [name,range] of [['profile_core',`A${row}:I${row}`],['profile_status',`N${row}:O${row}`],
                           ['profile_notes',`S${row}:T${row}`]]){
  const png=await wb.render({sheetName:'Профили',range,scale:1.5,format:'png'});
  await fs.writeFile(path.join(out,'previews',name+'.png'),new Uint8Array(await png.arrayBuffer()));
}
if(attempts)throw new Error('Unexpected network');
console.log(JSON.stringify({row,network_attempts:attempts}));
