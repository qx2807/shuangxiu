const XLSX = require('../lib/xlsx.full.min.js');
const fs = require('fs');
const path = require('path');

// Read the generated Excel template
const excelPath = path.join(__dirname, '..', '企业作息与评价数据表_模板.xlsx');
console.log('Testing reading excel file:', excelPath);

const fileBuf = fs.readFileSync(excelPath);
const wb = XLSX.read(fileBuf, { type: 'buffer' });

console.log('Sheet Names:', wb.SheetNames);
const sheetName = wb.SheetNames[0];
const rawJson = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });

console.log('Headers detected:', rawJson[0]);
console.log('Total rows:', rawJson.length);

function findColumnIndexes(headers) {
  const indexes = {
    name: -1,
    schedule: -1,
    hours: -1,
    industry: -1,
    reviews: []
  };

  headers.forEach((header, idx) => {
    const h = header.toLowerCase().replace(/[\s\r\n]/g, '');
    if (indexes.name === -1 && (h.includes('公司') || h.includes('企业') || h.includes('单位') || h === 'company' || h === 'name')) {
      indexes.name = idx;
      return;
    }
    if (indexes.schedule === -1 && (h.includes('作息') || h.includes('双休') || h.includes('制度') || h.includes('加班') || h === 'schedule')) {
      indexes.schedule = idx;
      return;
    }
    if (indexes.hours === -1 && (h.includes('时间') || h.includes('上下班') || h.includes('工时') || h === 'hours' || h === 'time')) {
      indexes.hours = idx;
      return;
    }
    if (indexes.industry === -1 && (h.includes('行业') || h.includes('类别') || h.includes('领域') || h === 'industry')) {
      indexes.industry = idx;
      return;
    }
    if (h.includes('评价') || h.includes('评论') || h.includes('反馈') || h.includes('口碑') || h.includes('备注') || h.includes('review') || h.includes('comment')) {
      indexes.reviews.push(idx);
      return;
    }
  });

  return indexes;
}

function splitMultiLineReviews(text) {
  if (!text) return [];
  let lines = text.split(/\r?\n/);
  if (lines.length === 1 && (text.includes('；') || text.includes(';'))) {
    lines = text.split(/[；;]/);
  }
  const cleanLines = [];
  lines.forEach(line => {
    let trimmed = line.trim();
    if (!trimmed) return;
    trimmed = trimmed.replace(/^(\d+[\.、\)]|\(\d+\)|[-*•·]\s*)/, '').trim();
    if (trimmed) cleanLines.push(trimmed);
  });
  return cleanLines;
}

const cols = findColumnIndexes(rawJson[0]);
console.log('Mapped Columns:', cols);

let parsedCount = 0;
let reviewTotal = 0;
for (let i = 1; i < rawJson.length; i++) {
  const row = rawJson[i];
  if (!row || !row[cols.name]) continue;
  parsedCount++;
  const revs = splitMultiLineReviews(row[cols.reviews[0]]);
  reviewTotal += revs.length;
}

console.log(`Test Result: Successfully parsed ${parsedCount} companies, total ${reviewTotal} reviews!`);
