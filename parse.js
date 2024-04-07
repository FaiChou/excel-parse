const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('data.xlsx');

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(worksheet);
console.log(data)

const resultJson = {};

for (const row of data) {
  const start = parseInt(row['A'], 10);
  const end = parseInt(row['B'], 10);
  const value = row['C'];
  for (let num = start; num <= end; num++) {
    resultJson[num.toString().padStart(5, '0')] = value;
  }
}

const findValue = (number) => {
  const numberStr = number.toString().padStart(5, '0');
  return resultJson[numberStr] || '未找到';
};

console.log(findValue(10017));  // 应输出：递送区域附加费-偏远
console.log(findValue(10001));  // 应输出：递送区域附加费

const jsonString = JSON.stringify(resultJson, null, 2);
fs.writeFileSync('result.json', jsonString);
