// Read-only artifact-tool preview of the changed technical heading.
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {createRequire} from 'node:module';
import net from 'node:net';
import http from 'node:http';
import https from 'node:https';
let requests=0;
function denied(){requests++;throw new Error('Offline preview');}
net.connect=net.createConnection=net.Socket.prototype.connect=denied;
http.get=http.request=https.get=https.request=globalThis.fetch=denied;
const out=path.dirname(fileURLToPath(import.meta.url));
const require=createRequire(path.join(out,'../../profiles_stage4/workbook.mjs'));
const {FileBlob,SpreadsheetFile}=await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const book=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(out,'verified_template.xlsx')));
const png=await book.render({sheetName:'Профили',range:'K1:P5',scale:1.5,format:'png'});
await fs.writeFile(path.join(out,'previews/technical_caption.png'),new Uint8Array(await png.arrayBuffer()));
if(requests)throw new Error('Unexpected network');
console.log(JSON.stringify({network_requests:requests,range:'K1:P5'}));
