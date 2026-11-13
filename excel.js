'use strict';

function exportRows(pkg, format) {
 const {values, refs, result} = pkg;
 const items = refs.units.filter(item => pkg.selected.includes(item.id));
 const client = format === 'klien';
 const offerColumn = client ? 0 : 5;
 const rows = [];
 const merges = [];
 function put(row, col, value, style = 0) {
  rows[row - 1] ??= [];
  rows[row - 1].push({col, value, style});
 }
 function offer(row, index, value, style) {
  put(row, offerColumn + index, value, style);
 }
 if (!client) {
  put(1, 0, values.name, 1);
  merges.push('A1:C1');
  put(2, 1, 'Jumlah orang');
  put(2, 2, result.people);
  ['No.', 'Parameter Pemeriksaan', 'Unit Cost'].forEach((label, col) => put(3, col, label, 1));
  items.forEach((item, index) => {
   put(4 + index, 0, index + 1);
   put(4 + index, 1, item.name);
   put(4 + index, 2, item.cost, 2);
  });
  const summary = [
   ['Operasional Cost', result.operating], ['Total Unit Cost', result.cost],
   ['Margin ' + values.margin + '%', result.margin], ['PPN ' + values.tax + '%', result.tax],
   ['Tarif Usulan', result.beforeFee], ['Referral Fee ' + values.fee + '%', result.fee],
   ['Tarif Jual', result.offer],
   [values.manual === '' ? 'Pembulatan' : 'Tarif Akhir Manual', result.final],
   ['Total Penawaran', result.total]
  ];
  summary.forEach(([label, value], index) => {
   const row = 4 + items.length + index;
   put(row, 1, label, index === 2 ? 6 : index >= 7 ? 8 : 0);
   put(row, 2, value, index === 2 ? 7 : index >= 7 ? 9 : 2);
  });
 }
 offer(1, 0, values.name, 1);
 merges.push(client ? 'A1:C1' : 'F1:H1');
 ['No.', 'Parameter Pemeriksaan', 'Tarif Penawaran / Orang'].forEach((label, col) => offer(3, col, label, 3));
 items.forEach((item, index) => {
  offer(4 + index, 0, index + 1, 5);
  offer(4 + index, 1, item.name, 4);
  offer(4 + index, 2, index === 0 ? result.final : '', 10);
 });
 const tariffCol = client ? 'C' : 'H';
 if (items.length > 1) merges.push(tariffCol + '4:' + tariffCol + (3 + items.length));
 return {rows, merges};
}

const xlsxNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function excelXml(value) {
 return String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g,'')
  .replace(/[<>&"']/g, character => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'})[character]);
}

function workbookParts({rows, merges}, format) {
 const client = format === 'klien';
 const columns = (client ? [8, 46, 25] : [8, 46, 18, 5, 5, 8, 46, 25]).map((width, i) =>
  '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+width+'" customWidth="1"/>').join('');
 const body = rows.map((row, index) => {
  const rowNumber = index + 1;
  const height = index === 0 ? 44 : index === 2 ? 27 : 25;
  return '<row r="'+rowNumber+'" ht="'+height+'" customHeight="1">'+row.map(({col, value, style}) => {
   const ref = String.fromCharCode(65+col)+rowNumber;
   if(typeof value === 'number') return '<c r="'+ref+'" s="'+style+'"><v>'+value+'</v></c>';
   return '<c r="'+ref+'" s="'+style+'" t="inlineStr"><is><t xml:space="preserve">'+excelXml(value)+'</t></is></c>';
  }).join('')+'</row>';
 }).join('');
 const mergeXml = '<mergeCells count="'+merges.length+'">'+merges.map(ref=>'<mergeCell ref="'+ref+'"/>').join('')+'</mergeCells>';
 const formats = [
  [0,0,0,0,'left'], [0,1,0,0,'center'], [3,0,0,0,'right'],
  [0,1,0,1,'center'], [0,0,0,1,'left'], [0,0,0,1,'center'],
  [0,0,2,0,'left'], [3,0,2,0,'right'], [0,1,0,0,'left'], [3,1,0,0,'right'], [3,1,0,1,'center']
 ];
 const xfs = formats.map(([numFmtId,fontId,fillId,borderId,horizontal]) =>
  `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="${horizontal}" vertical="center" wrapText="1"/></xf>`).join('');
 const styles = `<styleSheet xmlns="${xlsxNs}">
 <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
 <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>
 <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>
 <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
 <cellXfs count="${formats.length}">${xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
 </styleSheet>`;
 return {
  '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  '_rels/.rels':`<Relationships xmlns="${packageRel}"><Relationship Id="rId1" Type="${officeRel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  'xl/workbook.xml':`<workbook xmlns="${xlsxNs}" xmlns:r="${officeRel}"><sheets><sheet name="${client?'Penawaran MCU':'Perhitungan MCU'}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels':`<Relationships xmlns="${packageRel}"><Relationship Id="rId1" Type="${officeRel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${officeRel}/styles" Target="styles.xml"/></Relationships>`,
  'xl/styles.xml':styles,
  'xl/worksheets/sheet1.xml':`<worksheet xmlns="${xlsxNs}"><cols>${columns}</cols><sheetData>${body}</sheetData>${mergeXml}</worksheet>`
 };
}

const crcTable = Uint32Array.from({length:256}, (_, index) => {
 let value = index;
 for(let bit=0; bit<8; bit++) value = value & 1 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
 return value;
});

function excelZip(parts) {
 const encoder = new TextEncoder();
 const local = [], central = [];
 let offset = 0, centralSize = 0;
 function header(size, words, shorts) {
  const view = new DataView(new ArrayBuffer(size));
  for(const [at,value] of words) view.setUint32(at,value,true);
  for(const [at,value] of shorts) view.setUint16(at,value,true);
  return view.buffer;
 }
 for(const [path, xml] of Object.entries(parts)) {
  const name = encoder.encode(path);
  const bytes = encoder.encode(xmlHeader + xml);
  let crc = 0xffffffff;
  for(const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 255];
  crc = (crc ^ 0xffffffff) >>> 0;
  local.push(header(30,[[0,0x04034b50],[14,crc],[18,bytes.length],[22,bytes.length]],[[4,20],[12,33],[26,name.length]]),name,bytes);
  central.push(header(46,[[0,0x02014b50],[16,crc],[20,bytes.length],[24,bytes.length],[42,offset]],[[4,20],[6,20],[14,33],[28,name.length]]),name);
  offset += 30 + name.length + bytes.length;
  centralSize += 46 + name.length;
 }
 const count = Object.keys(parts).length;
 const end = header(22,[[0,0x06054b50],[12,centralSize],[16,offset]],[[8,count],[10,count]]);
 return new Blob([...local,...central,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

function downloadPackage(pkg, format) {
 const file = excelZip(workbookParts(exportRows(pkg,format),format));
 const slug = pkg.values.name.normalize('NFKC').replace(/[^\p{L}\p{N}]+/gu,'-').replace(/^-|-$/g,'').slice(0,70) || 'paket-mcu';
 const url = URL.createObjectURL(file);
 const link = document.createElement('a');
 link.href = url;
 link.download = slug + '-' + format + '.xlsx';
 document.body.append(link);
 link.click();
 link.remove();
 setTimeout(() => URL.revokeObjectURL(url),30000);
}
