const style = require('./style');


function createSludge(sheet,data){
  let index = sheet.lastRow.number;
  index++;
  sheet.mergeCells("A"+index,"P"+index);
  let titleCell = sheet.getCell("A"+index);
  titleCell.value = "七、泥区部分";
  const titleCol = sheet.getRow(index);
  titleCol.height = 20;
  Object.assign(titleCell,style.secondTitleStyle);
}


module.exports = createSludge;