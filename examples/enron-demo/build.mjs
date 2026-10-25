// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const outputDir = path.resolve(process.argv[2] ?? 'outputs/01a16b0a-0074-7383-ba33-19bcf4c27b0f');
const supportDir = path.resolve('artifacts/enron-demo');
await fs.mkdir(outputDir, {recursive:true});
await fs.mkdir(supportDir, {recursive:true});
const workbook = Workbook.create();
const research = workbook.worksheets.add('Research');
const queries = workbook.worksheets.add('Queries');
const p = {ink:'#273549',navy:'#28466B',muted:'#647184',pale:'#EDF2F7',input:'#FFF0B3',rule:'#CAD4E0'};
for (const [sheet,extent] of [[research,'A1:G39'],[queries,'A1:D29']]) {
  sheet.showGridLines = false;
  sheet.getRange(extent).format.font = {name:'Arial',size:11,color:p.ink};
  sheet.getRange(extent).format.rowHeight = 23;
  sheet.getRange(extent).format.verticalAlignment = 'center';
  sheet.getRange('A:A').format.columnWidth = 3;
  sheet.getRange('B2').format.font = {name:'Arial',size:17,bold:true,color:p.navy};
  sheet.getRange('B2').format.rowHeight = 32;
}
for (const [col,width] of Object.entries({B:24,C:47,D:3,E:46,F:16,G:3})) research.getRange(`${col}:${col}`).format.columnWidth = width;
for (const [col,width] of Object.entries({B:24,C:112,D:3})) queries.getRange(`${col}:${col}`).format.columnWidth = width;
research.tabColor = p.navy;
queries.tabColor = '#92A6BF';
const text = (s,c,v) => { s.getRange(c).values = [[v]]; };
const note = (s,c,v) => { text(s,c,v); s.getRange(c).format.font.color = p.muted; };
function header(s,range) {
  s.getRange(range).format = {fill:p.navy,font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center'};
}
text(research,'B2','Enron research');
research.getRange('B3:F3').format.borders = {bottom:{style:'thin',color:p.rule}};
research.getRange('B3').format.rowHeight = 8;
note(research,'B4','Hetzdev S3 source. Counts include mailbox copies. Dates and year boundaries use UTC.');
research.getRange('B5:B8').values = [['Parquet source'],['Year'],['Subject contains'],['Sender (exact)']];
research.getRange('C5:C8').values = [['s3://enron-emails/eml_messages.parquet'],[2001],['california'],['All']];
research.getRange('C5:C8').format.fill = p.input;
research.getRange('C6').setNumberFormat('0');
research.getRange('C6:C8').format.font.bold = true;
research.dataValidations.add({range:'C6',rule:{type:'whole',operator:'between',formula1:1900,formula2:2100}});
note(research,'B9','Amber cells are editable. Use All to remove either text filter.');
research.getRange('B10:B14').values = [['Loaded year'],['Filtered records'],['Monthly chart'],['Top senders chart'],['First 25 records']];
note(research,'B15','Select a handle above, then press Ctrl+Shift+P to open its preview.');
note(research,'E11','Loaded records');
note(research,'E13','Text filters reuse the loaded year.');
note(research,'E14','Changing the year reads S3 again.');
text(research,'B17','Monthly activity');
text(research,'E17','Top ten senders');
research.getRange('B17:F17').format.font.bold = true;
header(research,'B18:C18'); header(research,'E18:F18'); header(research,'E6:F6');
research.getRange('B19:B30').setNumberFormat('mmm yyyy');
research.getRange('C19:C30').setNumberFormat('#,##0');
research.getRange('F7:F8').setNumberFormat('#,##0');
research.getRange('F9').setNumberFormat('0.0%');
research.getRange('F11').setNumberFormat('#,##0');
research.getRange('F19:F28').setNumberFormat('#,##0');
for (const [row,message] of [
  [33,'Edit SQL on Queries. Each SQL cell is followed by the DuckQuery formula that runs it.'],
  [34,'Try energy as the subject filter, or copy a sender from the table into C8.'],
  [35,'The year table keeps five metadata columns. Email bodies stay in the source.'],
  [36,'Opening this workbook reloads data. F9 reuses unchanged results. Restart workbook refreshes S3.'],
  [37,'Connections are saved locally for this file. Keys are in Windows Credential Manager.'],
  [38,'A copy at a new path needs its connection set in XlDuck > Connections. Hetzdev must be reachable.']
]) note(research,`B${row}`,message);

text(queries,'B2','SQL and retained results');
note(queries,'B3','SQL is ordinary cell text. Parameters below bind values and table handles in question-mark order.');
const sql = {
  C5:`SELECT timezone('UTC', date) AS sent_at,
       lower(coalesce(nullif(trim(from_email), ''),
                      nullif(trim(from_name), ''), '(unknown)')) AS sender,
       coalesce(subject_clean, subject, '') AS subject,
       custodian, coalesce(is_reply, false) AS is_reply
FROM read_parquet(?)
WHERE date >= CAST(? AS TIMESTAMPTZ) AND date < CAST(? AS TIMESTAMPTZ)`,
  C9:`SELECT * FROM ?
WHERE (? = 'All' OR contains(lower(subject), lower(?)))
  AND (? = 'All' OR sender = lower(?))`,
  C13:`SELECT 'Matching records' AS Metric, count(*)::DOUBLE AS Value FROM ?
UNION ALL
SELECT 'Sender identifiers', count(DISTINCT sender)::DOUBLE FROM ?
UNION ALL
SELECT 'Reply share', avg(is_reply::INTEGER) FROM ?`,
  C17:`SELECT date_trunc('month', sent_at)::DATE AS Month,
       count(*) AS Records
FROM ? GROUP BY 1 ORDER BY 1`,
  C21:`SELECT sender AS Sender, count(*) AS Records
FROM ? GROUP BY 1
ORDER BY Records DESC, Sender LIMIT 10`,
  C25:`SELECT sent_at, sender, subject, custodian, is_reply
FROM ? ORDER BY sent_at, sender, subject, custodian
LIMIT 25`
};
const blocks = [[5,'1. Load a year','Only this query reads S3. There is no row cap on the selected year.'],[9,'2. Filter the year','Both filters use the same retained table. Subject matching is a literal, case-insensitive substring.'],[13,'3. Summarise','Counts include duplicates. Reply share uses the source is_reply flag; empty selections leave it blank.'],[17,'4. Monthly activity','Months with no matching records are omitted.'],[21,'5. Rank senders','Senders use email, then display name. These identifiers are not deduplicated people.'],[25,'6. Inspect records','This is the earliest 25 matching records, not a random sample. Times are UTC.']];
for (const [row,label,description] of blocks) {
  text(queries,`B${row}`,label);
  queries.getRange(`B${row}`).format.font.bold = true;
  text(queries,`C${row}`,sql[`C${row}`]);
  queries.getRange(`C${row}`).format.fill = '#F5F7FA';
  queries.getRange(`C${row}`).format.wrapText = true;
  queries.getRange(`C${row}`).format.verticalAlignment = 'top';
  queries.getRange(`C${row}`).format.rowHeight = (sql[`C${row}`].split('\n').length * 17) + 12;
  text(queries,`B${row+1}`,'Result handle');
  note(queries,`C${row+2}`,description);
  queries.getRange(`C${row+2}`).format.font.size = 10;
}
const formulas = {
  Queries: {
    C6:'=DuckQuery(C5,Research!C5,Research!C6&"-01-01 00:00:00+00",(Research!C6+1)&"-01-01 00:00:00+00")',
    C10:'=DuckQuery(C9,C6,Research!C7,Research!C7,Research!C8,Research!C8)',
    C14:'=DuckQuery(C13,C10,C10,C10)',
    C18:'=DuckQuery(C17,C10)',
    C22:'=DuckQuery(C21,C10)',
    C26:'=DuckQuery(C25,C10)'
  },
  Research: {
    C10:'=Queries!C6',C11:'=Queries!C10',
    C12:'=DuckPlot(Queries!C18,"line","x","Month","y","Records","title","Monthly mail records (UTC)")',
    C13:'=DuckPlot(Queries!C22,"bar","x","Sender","y","Records","title","Top senders by mail records")',
    C14:'=Queries!C26', E6:'=DuckOut(Queries!C14)', F11:'=DuckInspect(Queries!C6,"rows")',
    B18:'=DuckOut(Queries!C18)', E18:'=DuckOut(Queries!C22)'
  }
};
for (const [name,cells] of Object.entries(formulas)) for (const [cell,formula] of Object.entries(cells)) workbook.worksheets.getItem(name).getRange(cell).formulas = [[formula]];
for (const [sheet,cells] of [[research,['C10','C11','C12','C13','C14']],[queries,['C6','C10','C14','C18','C22','C26']]]) {
  for (const cell of cells) {
    sheet.getRange(cell).format.fill = p.pale;
    sheet.getRange(cell).format.font.color = p.navy;
    sheet.getRange(cell).setNumberFormat('General');
  }
}
workbook.recalculate();
const contract = {formulas,sql,source:'s3://enron-emails/eml_messages.parquet',ranges:{Research:'A1:G39',Queries:'A1:D29'}};
await fs.writeFile(path.join(supportDir,'contract.json'),JSON.stringify(contract,null,2));
const inspection = await workbook.inspect({kind:'table',range:'Research!B5:F11',include:'values,formulas',tableMaxRows:7,tableMaxCols:5,maxChars:2000});
await fs.writeFile(path.join(supportDir,'authoring-inspection.ndjson'),inspection.ndjson);
if (process.argv[3] === '--render-values') {
  // Native Excel supplies UDF values for visual review only. Never export this copy.
  const snapshot = JSON.parse((await fs.readFile(process.argv[4],'utf8')).replace(/^\uFEFF/,''));
  for (const [name,range] of Object.entries(contract.ranges)) workbook.worksheets.getItem(name).getRange(range).values = snapshot[name];
  for (const [name,range,file] of [['Research','B2:F30','research.png'],['Research','B33:F38','notes.png'],['Queries','B2:C15','queries-top.png'],['Queries','B17:C27','queries-bottom.png']]) {
    const img = await workbook.render({sheetName:name,range,scale:1.3,format:'png'});
    await fs.writeFile(path.join(supportDir,file),new Uint8Array(await img.arrayBuffer()));
  }
  console.log('Rendered native-calculated Enron demo.');
} else {
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const destination = path.join(outputDir,'XlDuck Enron Demo.xlsx');
  await output.save(destination);
  // The exporter writes a diagnostic sidecar; keep it with the authoring evidence.
  await fs.rename(destination + '.inspect.ndjson',path.join(supportDir,'export-inspection.ndjson')).catch(error => { if (error.code !== 'ENOENT') throw error; });
  console.log(destination);
}
