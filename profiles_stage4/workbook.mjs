// Value-only offline Excel renderer. This is the sole workbook authoring path.
import fs from 'node:fs/promises';
import http from 'node:http';
import https from 'node:https';
import net from 'node:net';
import dgram from 'node:dgram';
import dns from 'node:dns';
let networkAttempts = 0;
function denied() { networkAttempts++; throw new Error('Stage 4: network disabled'); }
http.request = http.get = https.request = https.get = denied;
net.connect = net.createConnection = net.Socket.prototype.connect = denied;
dgram.Socket.prototype.send = denied;
dns.lookup = dns.resolve = denied;
globalThis.fetch = denied;
const { Workbook, SpreadsheetFile } = await import('@oai/artifact-tool');
const [input, output, flag, previewDir] = process.argv.slice(2);
const payload = JSON.parse(await fs.readFile(input, 'utf8'));
const wb = Workbook.create();
function col(n) { let result=''; for(n++;n;n=Math.floor((n-1)/26)) result=String.fromCharCode(65+(n-1)%26)+result; return result; }
let index=0;
for (const [name, matrix] of Object.entries(payload.sheets)) {
  const sheet = wb.worksheets.add(name);
  const headers = matrix[0];
  // Excel table headings are unique without regard to case: r and R collide.
  if (headers.some(h=>typeof h!=='string' || !h.trim() || /[\x00-\x1f]/.test(h)) ||
      new Set(headers.map(h=>h.toLowerCase())).size!==headers.length) {
    throw new Error(`Invalid or case-insensitive duplicate table headers: ${name}`);
  }
  const dates = headers.map((s,i)=>s.startsWith('Дата')?i:-1).filter(i=>i>=0);
  const values = matrix.map((r,i)=>r.map((v,c)=>{
    if (i && dates.includes(c) && v) return new Date(v);
    if (typeof v==='string' && /^[=+@-]/.test(v)) return "'"+v;
    return v;
  }));
  const last = col(headers.length-1);
  const range = sheet.getRange(`A1:${last}${values.length}`);
  range.values = values;
  range.format.font = {name:'Arial', size:10, color:'#222222'};
  range.format.rowHeight = 24;
  range.format.verticalAlignment = 'center';
  range.format.wrapText = false;
  const head = sheet.getRange(`A1:${last}1`);
  head.format.fill = '#263E53';
  head.format.font = {name:'Arial',size:10,bold:true,color:'#FFFFFF'};
  head.format.horizontalAlignment='center';
  head.format.wrapText=true;
  head.format.rowHeight=42;
  for (let c=0;c<headers.length;c++) {
    const h=headers[c]; const column=sheet.getRange(`${col(c)}1:${col(c)}${values.length}`);
    const widths = {'Профиль':22,'Ключ':23,'Тип':31,'НД':25,'Источник':23,'Дата получения':22,'Статус':46,
      'УЗК правило':32,'Source ID':30,'Evidence reference':42,'Conflict ID':24,'Примечание':100,
      'Тип источника':45,'Family':31,'URL / файл':72,'Snapshot/cache':80,'SHA-256':68,'Количество записей':20,
      'Предпочтительный НД':27,'Основание':56,'Дата':22,'Источник A':35,'Источник B':35,
      'Значения A':105,'Значения B':105,'Поля конфликта':28,'Решение':32,'Дата решения':22,'Классификация':22,
      'Правило':33,'Часть':15,'L, мм':24,'S, мм':15,'Частота, МГц':34,'Угол, °':34,'Область':20};
    column.format.columnWidth=widths[h]??12;
    if (values.length>1) {
      const body=sheet.getRange(`${col(c)}2:${col(c)}${values.length}`);
      if (['h','b','b2','s','t','D','r','Радиус R','Количество записей'].includes(h)) {
        body.setNumberFormat(h==='Количество записей'?'0':'General'); body.format.horizontalAlignment='right';
      } else if (dates.includes(c)) { body.setNumberFormat('yyyy-mm-dd hh:mm'); }
      else { body.format.horizontalAlignment='left'; }
    }
  }
  if (values.length>1) {
    const table=sheet.tables.add(`A1:${last}${values.length}`,true,`ProfilesStage4Table${++index}`);
    table.showFilterButton=true;
    sheet.freezePanes.freezeRows(1);
    sheet.freezePanes.freezeColumns(name==='Профили'?2:1);
    const statusIndex=headers.indexOf('Статус');
    if(statusIndex>=0) sheet.getRange(`${col(statusIndex)}2:${col(statusIndex)}${values.length}`).conditionalFormats.add(
      'containsText',{text:'КОНФЛИКТ',format:{fill:'#FBE2DF',font:{color:'#A62121',bold:true}}});
    if(statusIndex>=0 && ['Профили','Конфликты'].includes(name)) sheet.getRange(`${col(statusIndex)}2:${col(statusIndex)}${values.length}`).conditionalFormats.add(
      'containsText',{text:name==='Конфликты'?'СПРАВОЧНЫЕ':'ЗАМЕЧАНИЯМИ',format:{fill:'#FFF0C2',font:{color:'#765000',bold:true}}});
  }
  if(name==='Правила УЗК') {
    sheet.getRange(`A2:I${values.length}`).format.wrapText=true;
    sheet.getRange(`A2:I${values.length}`).format.rowHeight=60;
  }
  if(name==='Конфликты' && values.length>1) {
    for(const c of ['F','H']) { sheet.getRange(`${c}2:${c}${values.length}`).format.wrapText=true; }
    sheet.getRange(`A2:M${values.length}`).format.rowHeight=44;
  }
  sheet.showGridLines=false;
  if(name==='Профили') sheet.tabColor='#263E53';
}
wb.recalculate();
const inspected=await wb.inspect({kind:'sheet',include:'name',maxChars:1800});
if (flag==='--preview') {
  await fs.mkdir(previewDir,{recursive:true});
  let i=0;
  for(const name of Object.keys(payload.sheets)) {
    const preview=await wb.render({sheetName:name,range:name==='Профили'?'A1:K10':name==='Источники'?'A1:E7':
      name==='Конфликты'?'I1:M6':name==='Предпочтения'?'A1:E3':'A1:I5',scale:1.5,format:'png'});
    await fs.writeFile(`${previewDir}/${++i}.png`,new Uint8Array(await preview.arrayBuffer()));
  }
  await fs.writeFile(`${previewDir}/inspect.ndjson`,inspected.ndjson);
}
if(networkAttempts) throw new Error('Network attempted during renderer');
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(output);
console.log(JSON.stringify({network_attempts:networkAttempts,sheets:Object.keys(payload.sheets),rows:payload.sheets['Профили'].length-1}));
