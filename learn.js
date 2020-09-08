const fs = require('fs')
const Excel = require('exceljs')

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

  // sheet1.autoFilter = 'A1:C1';
  // sheet1.getRow(5).font = {size:14,bold:true};
  // sheet1.getCell("A2").value = "Site";
  // sheet1.getCell("A2").font = {
  //   name:"Arial Black",
  //   color:{argb:"FF00FF00"},
  //   family:2,
  //   size:14,
  //   italic:true,
  //   bold:true
  // };

  // sheet1.columns = [{
  //   header:"Rating Period",
  //   key:"id",
  //   width:38
  // },{
  //   header:"Name",
  //   key:'name',
  //   width:32
  // },{
  //   header:"D.O.B.",
  //   key:"DOB",
  //   width:10,
  //   style:{
  //     numFmt:"dd/mm/yyyy"
  //   }
  // }];
  // sheet1.addRow({id:1,name:"John Doe",dob:new Date(1970,1,1)});
  // sheet1.addRow({id:2,name:"Jane Doe",dob:new Date(1965,1,7)});
  // sheet1.mergeCells("A4:A7");
  // sheet1.getCell("A6").value = '1989';

  // sheet1.mergeCells('A10',"B11");

  //设置格式化展示样式
  // sheet1.getCell("A1").value = 0.6;
  // sheet1.getCell("A1").numFmt = '# ?/?';

  //设置字体样式
  // let textRow = sheet1.getColumn("B");
  // textRow.width = 100;
  // let textCell = sheet1.getCell("B5");

  // textCell.value = "测试字体";
  // textCell.font = {
  //   name: "Comic Sans Ms",
  //   family: 4,
  //   size: 16,
  //   underline: true,
  //   bold: true
  // };
  // textCell.alignment = {
  //   vertical: "top",
  //   horizontal: "center"
  // }
  // textCell.border = {
  //   top: {
  //     style: "double",
  //     color: {
  //       argb: "FF00FF00"
  //     }
  //   },
  //   left: {
  //     style: "double"
  //   },
  //   bottom: {
  //     style: "thin"
  //   }
  // };
  // textCell.fill = {
  //   type: 'gradient',
  // gradient: 'path',
  // center:{left:0.5,top:0.5},
  // stops: [
  //   {position:0, color:{argb:'FFFF0000'}},
  //   {position:1, color:{argb:'FF00FF00'}}
  // ]
  // }


  workbook.xlsx.writeFile('./public/text.xlsx').then(res => {
    console.log('ready');
  })
}