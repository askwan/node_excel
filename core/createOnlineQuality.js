const style = require('./style');

const columns = [{
  start:(index)=>"A"+index,
  end:(index)=>"B"+index,
  name:"项目",
  key:"name"
},{
  start:(index)=>"C"+index,
  end:(index)=>"E"+index,
  name:"COD",
  key:"cod"
},{
  start:(index)=>"F"+index,
  end:(index)=>"H"+index,
  name:"NH3-N（mg/L）",
  key:"nh3_n"
},{
  start:(index)=>"I"+index,
  end:(index)=>"K"+index,
  name:"TP（mg/L）",
  key:"tp"
},{
  start:(index)=>"L"+index,
  end:(index)=>"N"+index,
  name:" TN（mg/L）",
  key:"tn"
},{
  start:(index)=>"O"+index,
  end:(index)=>"P"+index,
  name:"pH",
  key:"ph"
}]

function createOnlineQuality(sheet,data){
  let index = sheet.lastRow.number;
  index++;
  sheet.mergeCells("A"+index,"P"+index);
  let titleCell = sheet.getCell("A"+index);
  titleCell.value = "三、在线进出水水质";
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
  let onlineQuality = data.online_quality;
  onlineQuality.forEach(el=>{
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




module.exports = createOnlineQuality;