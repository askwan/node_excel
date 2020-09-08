const style = require('./style');

const columns = [{
  start:(index)=>"A"+index,
  end:(index)=>"B"+index,
  name:"项目",
  key:"name"
},{
  start:(index)=>"C"+index,
  end:(index)=>"D"+index,
  name:"BOD5      （mg/L）",
  key:"bod5"
},{
  start:(index)=>"E"+index,
  end:(index)=>"F"+index,
  name:"COD   （mg/L）",
  key:"cod"
},{
  start:(index)=>"G"+index,
  end:(index)=>"H"+index,
  name:"SS     （mg/L）",
  key:"ss"
},{
  start:(index)=>"I"+index,
  end:(index)=>"J"+index,
  name:"NH3-N（mg/L）",
  key:"nh3_n"
},{
  start:(index)=>"K"+index,
  end:(index)=>"K"+index,
  name:"TP   （mg/L）",
  key:"tp"
},{
  start:(index)=>"L"+index,
  end:(index)=>"L"+index,
  name:"TN   （mg/L）",
  key:"tn"
},{
  start:(index)=>"M"+index,
  end:(index)=>"M"+index,
  name:"pH",
  key:"ph"
},{
  start:(index)=>"N"+index,
  end:(index)=>"O"+index,
  name:"粪大肠菌群（CFU/L）",
  key:"cfu"
},{
  start:(index)=>"P"+index,
  end:(index)=>"P"+index,
  name:"色度(度)",
  key:"chroma"
}]

function createWaterQuality(sheet,data){
  let index = sheet.lastRow.number;
  index++;
  sheet.mergeCells("A"+index,"P"+index);
  let titleCell = sheet.getCell("A"+index);
  titleCell.value = "二、进出水水质";
  const titleCol = sheet.getRow(index);
  titleCol.height = 20;
  Object.assign(titleCell,style.secondTitleStyle);
  
  index++;
  columns.forEach(item=>{
    sheet.mergeCells(item.start(index), item.end(index));
    let cell = sheet.getCell(item.start(index));
    cell.border = style.border;
    cell.alignment = { ...style.alignCenter, ...style.wrap };
    cell.value = item.name;
    cell.font = style.fontStyle;
  });

  let waterQuality = data.water_quality;
  waterQuality.forEach(el=>{
    index++;
    columns.forEach(item=>{
      sheet.mergeCells(item.start(index), item.end(index));
      let cell = sheet.getCell(item.start(index));
      cell.value = el[item.key];
      cell.border = style.border;
      cell.alignment = { ...style.alignCenter, ...style.wrap };
      cell.font = style.fontStyle;
    })
  })

}

module.exports = createWaterQuality;