import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const root = process.argv[2];
const results = [];
for (const [name, rows] of [['creation', 3], ['conflict_True', 2]]) {
  const input = path.join(root, name, 'Документы', 'Гараж казахстан', 'Справочник марок.xlsx');
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
  const sheet = workbook.worksheets.getItemAt(0);
  const range = `A1:D${rows}`;
  const inspect = await workbook.inspect({ kind: 'table', range: `${sheet.name}!${range}`,
    include: 'values,formulas', tableMaxRows: rows, tableMaxCols: 4 });
  await fs.writeFile(path.join(root, `${name}_inspect.ndjson`), inspect.ndjson);
  const blob = await workbook.render({ sheetName: sheet.name, range, scale: 1.5, format: 'png' });
  const output = path.join(root, `${name}_catalog.png`);
  await fs.writeFile(output, new Uint8Array(await blob.arrayBuffer()));
  results.push({ input, image: output, range });
}
await fs.writeFile(path.join(root, 'render_results.json'), JSON.stringify(results, null, 2));
console.log(JSON.stringify(results));
