const style = require('./style');

const columns = [{
  start:index=>"A"+index,
  end:index=>"B"+(index+1),
  _start:index=>"A"+index,
  _end:index=>"B"+(index),
  name:"项目",
  key:"name"
},{
  start:index=>"C"+(index+1),
  end:index=>"D"+(index+1),
  _start:index=>"C"+index,
  _end:index=>"D"+(index),
  name:"回流量   (万m3)",
  key:"outer_flow"
},{
  start:index=>"E"+(index+1),
  end:index=>"F"+(index+1),
  _start:index=>"E"+index,
  _end:index=>"F"+(index),
  name:"回流比(%)",
  key:"outer_ratio"
},{
  start:index=>"G"+(index+1),
  end:index=>"H"+(index+1),
  _start:index=>"G"+index,
  _end:index=>"H"+(index),
  name:"MLSS   (mg/L)",
  key:"outer_mlss"
},{
  start:index=>"I"+(index+1),
  end:index=>"J"+(index+1),
  _start:index=>"I"+index,
  _end:index=>"J"+(index),
  name:"MLVSS  (mg/L)",
  key:"outer_mlvss"
},{
  start:index=>"K"+(index+1),
  end:index=>"L"+(index+1),
  _start:index=>"K"+index,
  _end:index=>"L"+(index),
  name:"回流量(万m3)",
  key:"inner_flow"
},{
  start:index=>"M"+(index+1),
  end:index=>"N"+(index+1),
  _start:index=>"M"+index,
  _end:index=>"N"+(index),
  name:"回流比(%)",
  key:"inner_ratio"
},{
  start:index=>"O"+index,
  end:index=>"P"+(index+1),
  _start:index=>"O"+index,
  _end:index=>"P"+(index),
  name:"污泥负荷(KgBOD5/Kg  MLSS.d)",
  key:"name"
}]

function createBackFlow(sheet,data){
  let index = sheet.lastRow.number;
  index++;
  sheet.mergeCells("A"+index,"P"+index);
  let titleCell = sheet.getCell("A"+index);
  titleCell.value = "六、回流部分";
  const titleCol = sheet.getRow(index);
  titleCol.height = 20;
  Object.assign(titleCell,style.secondTitleStyle);

  index++;
  sheet.mergeCells("C"+index,"J"+index);
  sheet.mergeCells("K"+index,"N"+index);
  let cell1 = sheet.getCell("C"+index);
  cell1.value = "外回流";
  cell1.border = style.border;
  cell1.alignment = { ...style.alignCenter, ...style.wrap };
  cell1.font = style.fontStyle;
  let cell2 = sheet.getCell("K"+index);
  cell2.value = "内回流";
  cell2.border = style.border;
  cell2.alignment = { ...style.alignCenter, ...style.wrap };
  cell2.font = style.fontStyle;
  columns.forEach(item=>{
    sheet.mergeCells(item.start(index), item.end(index));
    let cell = sheet.getCell(item.start(index));
    cell.border = style.border;
    cell.alignment = { ...style.alignCenter, ...style.wrap };
    cell.value = item.name;
    cell.font = style.fontStyle;
  });
  index++;

  
  let backflows = data.backflows;
  backflows.forEach(el=>{
    
    index++;
    columns.forEach(item=>{
      sheet.mergeCells(item._start(index), item._end(index));
      let cell = sheet.getCell(item._start(index));
      cell.value = el[item.key];
      cell.border = style.border;
      cell.alignment = { ...style.alignCenter, ...style.wrap };
      cell.font = style.fontStyle;
    })
  })


}

module.exports = createBackFlow;