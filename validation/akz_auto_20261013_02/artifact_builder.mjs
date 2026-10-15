import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const out=process.argv[2], mode=process.argv[3];
const setup=JSON.parse(await fs.readFile(path.join(out,'SETUP.json'),'utf8'));
const hash=bytes=>crypto.createHash('sha256').update(bytes).digest('hex');
try{
if(mode==='fixture'||mode==='fixture_r2'){
 const target=path.join(setup.workspace,'DEV','Stage6_Fixtures','assignments.xlsx');
 const original=await fs.readFile(target);const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(target));
 const s=wb.worksheets.getItemAt(0), range=s.getRange('A1:C196'), values=range.values;
 const assignment={'Бп2-23':0,'Бп2-5':1,'Бп2-18':2,'Бп2-19':3};const found=[];
 for(let r=1;r<values.length;r++){
   const name=values[r][0];if(Object.hasOwn(assignment,name)){s.getRange('C'+(r+1)).values=[[assignment[name]]];found.push(name);}
   else if(values[r][2]!==null&&values[r][2]!==undefined&&values[r][2]!==''){s.getRange('C'+(r+1)).values=[[null]];}
 }
 if(found.length!==4)throw new Error('TEST_ASSIGNMENT_MISSING '+JSON.stringify(found));
 wb.recalculate();const exported=await SpreadsheetFile.exportXlsx(wb);const temp=target+'.fixture.xlsx';await exported.save(temp);await fs.rename(temp,target);
 await fs.writeFile(path.join(out,mode==='fixture_r2'?'TEST_ASSIGNMENT_FIXTURE_R2.json':'TEST_ASSIGNMENT_FIXTURE.json'),JSON.stringify({target,before_sha256:hash(original),after_sha256:hash(await fs.readFile(target)),assignment,mode:'ARTIFICIAL_TEST_ASSIGNMENTS_ONLY',original_source_unchanged:true},null,2),{flag:'wx'});
 console.log('Isolated assignment fixture PASS');
}else if(mode==='before'||mode==='after'||mode==='after_exe'){
 const input=mode==='before'?'TEMPLATE_RENDER_INPUTS.json':mode==='after_exe'?'OUTPUT_EXE_RENDER_INPUTS.json':'OUTPUT_RENDER_INPUTS.json';
 const items=JSON.parse(await fs.readFile(path.join(out,input),'utf8'));
 const dir=path.join(out,'renders_'+mode);await fs.mkdir(dir,{recursive:true});
 const evidence=[];
 for(const item of items){
   const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(item.path));
   const view=await wb.inspect({kind:'table',range:`'${item.sheet}'!${item.range}`,include:'values,formulas',tableMaxRows:8,tableMaxCols:10,maxChars:1600});
   const image=await wb.render({sheetName:item.sheet,range:item.range,scale:1,format:'png'});
   const target=path.join(dir,item.name+'.png');await fs.writeFile(target,new Uint8Array(await image.arrayBuffer()));
   evidence.push({input:item.path,sha256:hash(await fs.readFile(item.path)),render:target,inspection:view.ndjson});
   console.log(mode,item.name,'rendered');
 }
 await fs.writeFile(path.join(out,'ARTIFACT_RENDER_'+mode.toUpperCase()+'.json'),JSON.stringify(evidence,null,2),{flag:'wx'});
}else throw new Error('Unknown mode');
}catch(error){console.error('ARTIFACT_ERROR:',error.message);process.exitCode=1;}
