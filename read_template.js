const xlsx = require('xlsx');
const wb = xlsx.readFile('../Escola Municipal Delphina Aziz - Iranduba - AM.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
console.log(JSON.stringify(xlsx.utils.sheet_to_json(ws, {header: 1})));
