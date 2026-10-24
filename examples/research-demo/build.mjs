// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const outputDir = path.resolve(process.argv[2] ?? 'outputs/01a16b0a-0074-7383-ba33-19bcf4c27b0f');
const supportDir = path.resolve('artifacts/demo-authoring');
await fs.mkdir(outputDir, {recursive:true});
await fs.mkdir(supportDir, {recursive:true});
const workbook = Workbook.create();
const demo = workbook.worksheets.add('Research');
const data = workbook.worksheets.add('Observations');
const palette = {ink:'#243247',muted:'#5F6B7A',navy:'#28466B',pale:'#EDF2F7',input:'#FFF0B3',purple:'#EEE9F8',rule:'#CAD4E0'};
for (const [sheet, extent] of [[demo,'A1:H47'],[data,'A1:G92']]) {
  sheet.showGridLines = false;
  sheet.getRange(extent).format.font = {name:'Arial',size:11,color:palette.ink};
  sheet.getRange(extent).format.rowHeight = 21;
  sheet.getRange(extent).format.verticalAlignment = 'center';
  sheet.getRange('A:A').format.columnWidth = 3;
}
for (const [column,width] of Object.entries({B:25,C:26,D:18,E:25,F:17,G:20,H:3}))
  demo.getRange(`${column}:${column}`).format.columnWidth = width;
for (const [column,width] of Object.entries({B:20,C:12,D:18,E:20,F:18,G:3}))
  data.getRange(`${column}:${column}`).format.columnWidth = width;
demo.tabColor = palette.navy;
data.freezePanes.freezeRows(6);
function text(sheet,address,value) { sheet.getRange(address).values = [[value]]; }
function title(sheet,title,end) {
  text(sheet,'B2',title);
  sheet.getRange(`B2:${end}2`).format.font = {name:'Arial',size:17,bold:true,color:palette.navy};
  sheet.getRange('B2').format.rowHeight = 30;
  sheet.getRange(`B3:${end}3`).format.borders = {bottom:{style:'thin',color:palette.rule}};
  sheet.getRange('B3').format.rowHeight = 8;
}
function section(sheet,range,label) {
  const r = sheet.getRange(range);
  r.format.fill = palette.pale;
  r.format.font = {name:'Arial',size:11,bold:true,color:palette.navy};
  text(sheet,range.split(':')[0],label);
}
title(demo,'XlDuck research playground','G');
text(demo,'B4','Enable this workbook from XlDuck > Enable / Connections.');
text(demo,'B5','Synthetic sensor data. Edit the amber inputs, then select a chart handle to preview it.');
demo.getRange('B4:G5').format.font.color = palette.muted;
demo.getRange('B7:B9').values = [['First day (1–28)'],['Site'],['Minimum response (a.u.)']];
demo.getRange('C7:C9').values = [[8],['All'],[35]];
demo.getRange('C7:C9').format.fill = palette.input;
demo.getRange('C7:C9').format.font.bold = true;
demo.getRange('C7').setNumberFormat('0');
demo.getRange('C9').setNumberFormat('0.0');
demo.dataValidations.add({range:'C7',rule:{type:'whole',operator:'between',formula1:1,formula2:28}});
demo.getRange('C8').dataValidation = {rule:{type:'list',values:['All','Estuary','Meadow','Ridge']}};
demo.dataValidations.add({range:'C9',rule:{type:'decimal',operator:'between',formula1:0,formula2:100}});
text(demo,'E7','Retained observations');
text(demo,'E8','Mean response (a.u.)');
text(demo,'E9','Blank mean = no matching observations.');
demo.getRange('E9:G9').format.font = {name:'Arial',size:10,italic:true,color:palette.muted};
demo.getRange('F7:F8').format.font.bold = true;
demo.getRange('F7').setNumberFormat('#,##0');
demo.getRange('F8').setNumberFormat('0.00');
text(demo,'B11','Daily mean response');
text(demo,'B12','Compare sites');
text(demo,'E11','Select C11, then press Ctrl+Shift+P.');
text(demo,'E12','Keep a chart selected while editing inputs.');
demo.getRange('E11:G12').format.font.color = palette.muted;
section(demo,'B14:C14','The query graph');
section(demo,'E14:G14','Current site summary');
demo.getRange('B15:C15').values = [['Step','Result handle']];
demo.getRange('B15:C15').format.font.bold = true;
demo.getRange('B16:B21').values = [['1. Capture observations'],['2. Filter day and site'],['3. Define response recipe'],['4. Materialise shared data'],['5. Summarise by site'],['6. Summarise by day']];
const formulas = {
  C16:'=DuckCapture(DemoObservations[#All])',
  C17:'=DuckQuery("SELECT * FROM ? WHERE accepted AND day >= ? AND (? = \'All\' OR site = ?) ORDER BY day, site",C16,$C$7,$C$8,$C$8)',
  C18:'=DuckFrag("SELECT * FROM ? WHERE response >= ?",C17,$C$9)',
  C19:'=DuckQuery("SELECT * FROM ?",C18)',
  C20:'=DuckQuery("SELECT site, count(*) AS samples, avg(response) AS mean_response FROM ? GROUP BY site ORDER BY site",C19)',
  C21:'=DuckQuery("SELECT day, avg(response) AS mean_response FROM ? GROUP BY day ORDER BY day",C19)',
  C11:'=DuckPlot(C21,"line","x","day","y","mean_response","title","Daily mean response (a.u.)")',
  C12:'=DuckPlot(C20,"bar","x","site","y","mean_response","title","Mean response by site (a.u.)")',
  F7:'=DuckInspect(C19,"rows")',
  F8:'=DuckQueryOutScalar("SELECT avg(response) FROM ?",C19)',
  E15:'=DuckOut(C20)',
  E22:'=DuckInspect(C19,"schema")',
  B29:'=DuckQueryOut("SELECT * FROM ? ORDER BY day, site LIMIT 8",C19)'
};
for (const [cell,formula] of Object.entries(formulas)) demo.getRange(cell).formulas = [[formula]];
const handleLabels = {C11:'Select for line chart',C12:'Select for bar chart',C16:'Captured table',C17:'Filtered table',C18:'Lazy SQL recipe',C19:'Shared table',C20:'Site summary',C21:'Daily means'};
for (const [cell,label] of Object.entries(handleLabels)) {
  demo.getRange(cell).setNumberFormat(`General;General;General;"${label}"`);
  demo.getRange(cell).format.fill = cell === 'C18' ? palette.purple : palette.pale;
  demo.getRange(cell).format.font.color = palette.navy;
}
demo.getRange('C11:C12').format.font.bold = true;
demo.getRange('C16:C21').format.rowHeight = 24;
demo.getRange('E15:G15').format = {fill:palette.navy,font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'}};
demo.getRange('E15:G15').format.horizontalAlignment = 'center';
demo.getRange('F16:F18').setNumberFormat('#,##0');
demo.getRange('G16:G18').setNumberFormat('0.00');
section(demo,'E21:F21','Filtered data schema');
demo.getRange('E22:F22').format.font.bold = true;
text(demo,'B23','Select a handle to inspect its data or edit its formula.');
text(demo,'B24','C18 is a recipe. C19 runs it once for the branches.');
demo.getRange('B23:G24').format.font = {name:'Arial',size:10,color:palette.muted};
section(demo,'B28:F28','First eight retained observations');
demo.getRange('B29:F29').format = {fill:palette.navy,font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'}};
demo.getRange('B29:F29').format.horizontalAlignment = 'center';
demo.getRange('B30:B37').setNumberFormat('@');
demo.getRange('C30:C37').setNumberFormat('0');
demo.getRange('E30:E37').setNumberFormat('0.00');
section(demo,'B39:G39','Try these edits');
for (const [row,line] of [
  [40,'Change C7 from 8 to 15, or choose a site in C8. The summary and selected chart update.'],
  [41,'Raise C9 to 60. Compare the retained sample count and mean response.'],
  [42,'Edit a response or accepted flag on Observations. The capture and downstream graph update.'],
  [43,'F9 and unchanged F2+Enter reuse results. Restart workbook creates a fresh session.'],
  [44,'Ask your Excel agent: compare site variability and add a point chart using this graph.'],
  [45,'Use Retry failed after fixing an error. #N/A can appear briefly while a query is running.']
]) text(demo,`B${row}`,line);

title(data,'Synthetic observations','F');
text(data,'B4','Source: generated demo data for three fictional sites over 28 days.');
text(data,'B5','Response is in arbitrary units. Only rows marked accepted enter the analysis.');
data.getRange('B4:F5').format.font = {name:'Arial',size:10,color:palette.muted};
const sites = ['Estuary','Meadow','Ridge'];
const rows = [];
for (let day=1;day<=28;day++) for (let site=0;site<3;site++) {
  const response = Math.round((38 + 9*site + 0.65*day + 3*Math.sin(day*.4+site) + (day%5)*.4)*100)/100;
  rows.push([String(rows.length+1).padStart(4,'0'),day,sites[site],response,(day+site*3)%13 !== 0]);
}
data.getRange('B6:F90').values = [['sample_id','day','site','response','accepted'],...rows];
data.getRange('B7:B90').setNumberFormat('@');
data.getRange('C7:C90').setNumberFormat('0');
data.getRange('E7:E90').setNumberFormat('0.00');
const table = data.tables.add('B6:F90',true,'DemoObservations');
table.style = 'TableStyleMedium2';
data.getRange('B6:F6').format.font = {name:'Arial',size:11,bold:true,color:'#FFFFFF'};
data.getRange('B6:F6').format.fill = palette.navy;
data.getRange('B6:F6').format.horizontalAlignment = 'center';
data.getRange('E7:F90').format.fill = '#FFF8DE';

workbook.recalculate();
await fs.writeFile(path.join(supportDir,'contract.json'),JSON.stringify({formulas,rows,handleLabels},null,2));
const inspection = await workbook.inspect({kind:'table',range:'Research!B7:G9',include:'values,formulas',tableMaxRows:3,tableMaxCols:6,maxChars:1500});
await fs.writeFile(path.join(supportDir,'authoring-inspection.ndjson'),inspection.ndjson);
if (process.argv[3] === '--render-values') {
  // The rendering library cannot execute Excel-DNA UDFs. Use values captured from
  // the verified native workbook for review only; never export this render copy.
  const snapshot = JSON.parse((await fs.readFile(process.argv[4],'utf8')).replace(/^\uFEFF/,''));
  demo.getRange('B1:G47').values = snapshot.research;
  for (const [cell,label] of Object.entries(handleLabels)) text(demo,cell,label);
  // The preview formatter normalizes numeric-looking text under @. The native
  // workbook tests verify the actual text and leading zeros; mirror that display.
  demo.getRange('B30:B37').setNumberFormat('0000');
  data.getRange('B7:B90').setNumberFormat('0000');
  for (const [sheet,range,file] of [[demo,'B2:G27','research-top.png'],[demo,'B28:G45','research-detail.png'],[data,'B2:F21','observations.png']]) {
    const image = await workbook.render({sheetName:sheet.name,range,scale:1.5,format:'png'});
    await fs.writeFile(path.join(supportDir,file),new Uint8Array(await image.arrayBuffer()));
  }
  console.log('Rendered native-calculated values for visual review.');
} else {
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  const filename = path.join(outputDir,'XlDuck Research Demo.xlsx');
  await xlsx.save(filename);
  const inspectPath = `${filename}.inspect.ndjson`;
  try { await fs.copyFile(inspectPath,path.join(supportDir,'artifact-inspect.ndjson')); await fs.unlink(inspectPath); }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  console.log(filename);
}
