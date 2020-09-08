
const Excel = require('exceljs');
const style = require('./core/style')

const createBaseInfo = require("./core/createBaseInfo");
const createWaterQuality = require('./core/createWaterQuality');
const createOnlineQuality = require('./core/createOnlineQuality');
const createWaterArea = require('./core/createWaterArea');
const createBackFlow = require('./core/createBackFlow');


const colstring = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const colIndex = colstring.split("");

const workbook = new Excel.Workbook();

workbook.creator = "askwan";
workbook.lastModifiedBy = "askwan";
workbook.created = new Date();
workbook.modified = new Date();
workbook.lastPrinted = new Date();

module.exports = function (data) {
  workbook.views = [
    {
      x: 0, y: 0, width: 1000, height: 2000,
      firstSheet: 0, activeTab: 2, visibility: "visible"
    }
  ];

  const sheet1 = workbook.addWorksheet('my sheet', { properties: { tabColor: { argb: "FFC0000" } } });
  colIndex.forEach(el=>{
    sheet1.getColumn(el).width = 5;
  })
  

  createTitle(sheet1,data);
  createBaseInfo(sheet1,data);
  createWaterQuality(sheet1,data);
  createOnlineQuality(sheet1,data);
  createWaterArea(sheet1,data);
  createBackFlow(sheet1,data);

  workbook.xlsx.writeFile('./public/test.xlsx').then(res => {
    console.log('ready');
  });
}

function createTitle(sheet,data){
  let index = 0;
  index++;
  const titleCell = sheet.getCell(colIndex[0]+index)
  sheet.mergeCells(`A${index}:P${index}`)
  titleCell.value = data.project_name;
  // titleCell.font = style.titleFontStyle;
  // titleCell.alignment = style.alignCenter;
  // titleCell.fill = style.titleFill;
  Object.assign(titleCell,style.titleStyle);
  index++;
  sheet.getCell("K"+index).value="编号：";
  sheet.getCell("L"+index).value = data.serial;
  return index;
}