const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat,
} = require('docx');

// ---------- Excel okuma / yazma ----------

function readExcel(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: false });
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const data = [];
    if (ws['!ref']) {
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = range.s.r; r <= range.e.r; r++) {
        const row = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell) { row.push(''); continue; }
          if (cell.f) row.push('=' + cell.f);
          else if (cell.w !== undefined) row.push(String(cell.w));
          else if (cell.v !== undefined && cell.v !== null) row.push(String(cell.v));
          else row.push('');
        }
        data.push(row);
      }
    }
    return { name, data };
  });
  return sheets;
}

function writeExcel(filePath, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const aoa = sheet.data.map((row) => row.map((v) => {
      if (v === '' || v === null || v === undefined) return null;
      const s = String(v);
      if (s.startsWith('=') && s.length > 1) return { t: 'n', v: 0, f: s.slice(1) }; // formül
      // Sayı mı? (Türkçe virgül de kabul et)
      const normalized = s.replace(/\s/g, '').replace(',', '.');
      if (/^-?\d+(\.\d+)?$/.test(normalized)) return Number(normalized);
      return s;
    }));
    const ws = XLSX.utils.aoa_to_sheet(aoa.length ? aoa : [[]]);
    const safeName = (sheet.name || 'Sayfa').slice(0, 31).replace(/[\\/?*[\]:]/g, '_');
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });
  wb.Workbook = { CalcPr: { fullCalcOnLoad: true } }; // Excel açılışta formülleri hesaplasın
  const ext = path.extname(filePath).toLowerCase();
  const bookType = ext === '.csv' ? 'csv' : ext === '.xls' ? 'biff8' : 'xlsx';
  XLSX.writeFile(wb, filePath, { bookType });
}

// ---------- Word okuma / yazma ----------

async function readWord(filePath) {
  const result = await mammoth.convertToHtml({ path: filePath }, { styleMap: ['u => u', 'strike => s'] });
  return result.value;
}

function buildDocx(blocks) {
  const children = blocks.map((b) => {
    const runs = (b.runs || []).map((r) => new TextRun({
      text: r.text || '',
      bold: !!r.bold,
      italics: !!r.italic,
      underline: r.underline ? {} : undefined,
      strike: !!r.strike,
    }));
    const opts = { children: runs.length ? runs : [new TextRun('')] };
    if (b.heading === 1) opts.heading = HeadingLevel.HEADING_1;
    else if (b.heading === 2) opts.heading = HeadingLevel.HEADING_2;
    else if (b.heading === 3) opts.heading = HeadingLevel.HEADING_3;
    if (b.align === 'center') opts.alignment = AlignmentType.CENTER;
    else if (b.align === 'right') opts.alignment = AlignmentType.RIGHT;
    else if (b.align === 'justify') opts.alignment = AlignmentType.JUSTIFIED;
    if (b.list === 'bullet') opts.bullet = { level: 0 };
    else if (b.list === 'number') opts.numbering = { reference: 'numaralar', level: 0 };
    return new Paragraph(opts);
  });

  return new Document({
    numbering: {
      config: [{
        reference: 'numaralar',
        levels: [{
          level: 0,
          format: LevelFormat.DECIMAL,
          text: '%1.',
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      }],
    },
    sections: [{ children: children.length ? children : [new Paragraph('')] }],
  });
}

async function writeWord(filePath, blocks) {
  const doc = buildDocx(blocks);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buffer);
}

module.exports = { readExcel, writeExcel, readWord, writeWord, buildDocx };
