const style = require('./style');

const columns = [{
  start:index=>"A"+index,
  end:index=>"A"+(index+1),
  _start:index=>"A"+index,
  _end:index=>"A"+(index),
  name:"项目",
  key:"name"
},{
  start:index=>"B"+index,
  end:index=>"C"+(index+1),
  _start:index=>"B"+index,
  _end:index=>"C"+(index),
  name:"进水量  (m3)",
  key:"inner_water"
},{
  start:index=>"D"+index,
  end:index=>"D"+(index+1),
  _start:index=>"D"+index,
  _end:index=>"D"+(index),
  name:"初沉量（m3）",
  key:"primary_sludge"
},{
  start:index=>"E"+index,
  end:index=>"E"+(index+1),
  _start:index=>"E"+index,
  _end:index=>"E"+(index),
  name:"剩余量（m3）",
  key:"last_sludge"
},{
  start:index=>"F"+(index+1),
  end:index=>"G"+(index+1),
  _start:index=>"F"+(index),
  _end:index=>"G"+(index),
  name:"供气量(m3)",
  key:"inner_gas"
},{
  start:index=>"H"+(index+1),
  end:index=>"H"+(index+1),
  _start:index=>"H"+(index),
  _end:index=>"H"+(index),
  name:"气水比",
  key:"gas_vs_water"
},{
  start:index=>"I"+(index+1),
  end:index=>"J"+(index+1),
  _start:index=>"I"+(index),
  _end:index=>"J"+(index),
  name:"MLSS  (mg/L)",
  key:"mlss"
},{
  start:index=>"K"+(index+1),
  end:index=>"L"+(index+1),
  _start:index=>"K"+(index),
  _end:index=>"L"+(index),
  name:"MLVSS (mg/L)",
  key:"mlvss"
},{
  start:index=>"M"+(index+1),
  end:index=>"N"+(index+1),
  _start:index=>"M"+(index),
  _end:index=>"N"+(index),
  name:"SV(%)",
  key:"sv"
},{
  start:index=>"O"+(index+1),
  end:index=>"P"+(index+1),
  _start:index=>"O"+(index),
  _end:index=>"P"+(index),
  name:"SVI  (mL/g)",
  key:"svi"
}]

function createWaterArea(sheet,data){
  let index = sheet.lastRow.number;
  index++;
  sheet.mergeCells("A"+index,"P"+index);
  let titleCell = sheet.getCell("A"+index);
  titleCell.value = "五、水区处理部分";
  const titleCol = sheet.getRow(index);
  titleCol.height = 20;
  Object.assign(titleCell,style.secondTitleStyle);

  index++;
  sheet.mergeCells("F"+index,"P"+index);
  let biocell = sheet.getCell("F"+index);
  biocell.value = "生物池";
  biocell.border = style.border;
  biocell.alignment = { ...style.alignCenter, ...style.wrap };
  biocell.font = style.fontStyle;
  columns.forEach(item=>{
    sheet.mergeCells(item.start(index), item.end(index));
    let cell = sheet.getCell(item.start(index));
    cell.border = style.border;
    cell.alignment = { ...style.alignCenter, ...style.wrap };
    cell.value = item.name;
    cell.font = style.fontStyle;
  });
  index++;

  
  let water_area = data.water_area;
  water_area.forEach(el=>{
    
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

module.exports = createWaterArea