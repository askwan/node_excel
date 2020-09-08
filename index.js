const fs = require('fs')

const createExcel = require('./createExcel');

// const data = require('data.json');
fs.readFile('data.json','utf8',(err,data)=>{
  let json = JSON.parse(data);
  createExcel(json.data.list[0]);
})
