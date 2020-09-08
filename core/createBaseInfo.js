const style = require('./style');

const titles = [{
  start: (index) => "A" + index,
  end: (index) => "A" + (index),
  value: "项目",
  key: "name"
}, {
  start: (index) => "B" + index,
  end: (index) => "D" + (index),
  value: "处理水量         (万m³)",
  key: "water"
}, {
  start: (index) => "E" + index,
  end: (index) => "G" + (index),
  value: "耗电量        (万kW·h)",
  key: "electric"
}, {
  start: (index) => "H" + index,
  end: (index) => "J" + (index),
  value: "电量单耗       (kW·h/ m3）",
  key: "electric_p"
}, {
  start: (index) => "K" + index,
  end: (index) => "M" + (index),
  value: "渣  量      （m3）",
  key: "dross",
}, {
  start: (index) => "N" + index,
  end: (index) => "P" + (index),
  value: "砂  量      （m3）",
  key: "sand"
}];

function createBaseInfo(sheet, data) {
  let index = sheet.lastRow.number;
  index++;
  sheet.getCell("A" + index).value = "一、进水量、电量及栅渣量";
  sheet.mergeCells(`A${index}:E${index}`);
  sheet.getCell("F" + index).value = data.create_time;
  sheet.getCell("J" + index).value = "星期：";
  sheet.getCell("K" + index).value = "周日";
  sheet.getCell("L" + index).value = "气温：";
  sheet.getCell("M" + index).value = data.air_templature + "°C";
  sheet.getCell("N" + index).value = "水温：";
  sheet.getCell("O" + index).value = data.water_templature + "°C";

  sheet.getRow(index).font = style.fontStyle;
  


  index++;
  
  titles.forEach(item => {
    sheet.mergeCells(item.start(index), item.end(index));
    sheet.getCell(item.start(index)).value = item.value;
  });
  let row = sheet.getRow(index);
  // row.height = 28;
  sheet.getRow(index).eachCell((cell) => {
    cell.border = style.border;
    cell.alignment = { ...style.alignCenter, ...style.wrap };
    cell.font = style.fontStyle;
  });

  // 数据
  let surveys = data.surveys;

  surveys.forEach((el, i) => {
    index ++;
    titles.forEach((item) => {
      sheet.mergeCells(item.start(index), item.end(index));
      sheet.getCell(item.start(index)).value = el[item.key];
      sheet.getRow(index).eachCell(cell => {
        cell.border = style.border;
        cell.alignment = { ...style.alignCenter, ...style.wrap };
        cell.font = style.fontStyle;
      })
    })
  });
  // 合计
  index++;
  titles.forEach(item=>{
    sheet.mergeCells(item.start(index),item.end(index));
    sheet.getCell(item.start(index)).value = sum(surveys,item.key);
    sheet.getRow(index).eachCell(cell => {
      cell.border = style.border;
      cell.alignment = { ...style.alignCenter, ...style.wrap };
      cell.font = style.fontStyle;
    })
  })

};

function sum(list,key){
  if(key === 'name') return "项目";
  const scale = 100000;
  let result = 0;
  list.forEach(el=>{
    result += Number(el[key])*scale;
  });
  return result/scale;
}

module.exports = createBaseInfo;