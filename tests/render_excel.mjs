import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const [input, range, output] = process.argv.slice(2);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const sheet = workbook.worksheets.getItemAt(0);
const preview = await workbook.render({ sheetName: sheet.name, range, scale: 1.5, format: 'png' });
await fs.writeFile(output, new Uint8Array(await preview.arrayBuffer()));
