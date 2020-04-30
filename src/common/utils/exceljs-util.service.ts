import { Injectable } from '@angular/core';
import { Xlsx } from 'exceljs';
import saveAs from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class ExceljsUtilService {

  Excel = require("exceljs");
  constructor() { }

  async test() {
    var workbook = new this.Excel.Workbook();
    workbook.creator = "Me";
    workbook.lastModifiedBy = "Her";
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();

    let sheet = workbook.addWorksheet('2018-10报表');
    sheet.columns = [
      { header: '创建日期', key: 'create_time', width: 15 },
      { header: '单号', key: 'id', width: 15 },
      { header: '电话号码', key: 'phone', width: 15 },
      { header: '地址', key: 'address', width: 15 }
    ];
    const data = [{
      create_time: '2018-10-01',
      id: '787818992109210',
      phone: '13611010535',
      address: '深圳市'
    }];

    sheet.addRows(data);

    // 添加表头
    sheet.getRow(1).values = ['种类', '销量', , , , '店铺'];
    sheet.getRow(2).values = ['种类', '2018-05', '2018-06', '2018-07', '2018-08', '店铺'];

    // 添加数据项定义，与之前不同的是，此时去除header字段
    sheet.columns = [
      { key: 'category', width: 30 },
      { key: '2018-05', width: 30 },
      { key: '2018-06', width: 30 },
      { key: '2018-07', width: 30 },
      { key: '2018-08', width: 30 },
      { key: 'store', width: 30 },
    ];
    const data2 = [{
      category: '衣服',
      '2018-05': 300,
      '2018-06': 230,
      '2018-07': 730,
      '2018-08': 630,
      'store': '王小二旗舰店'
    }, {
      category: '零食',
      '2018-05': 672,
      '2018-06': 826,
      '2018-07': 302,
      '2018-08': 389,
      'store': '吃吃货'
    }];
    sheet.addRows(data2);

    // 合并单元格
    sheet.mergeCells(`B1:E1`);
    sheet.mergeCells('A1:A2');
    sheet.mergeCells('F1:F2');
    sheet.getCell('A2').fill = {
      type: 'pattern',
      pattern: 'darkTrellis',
      fgColor: { argb: 'FFFFFF00' },
      bgColor: { argb: 'FF0000FF' }
    };


    // 设置每一列样式
    const row = sheet.getRow(1)
    row.eachCell((cell, rowNumber) => {
      sheet.getColumn(rowNumber).alignment = { vertical: 'middle', horizontal: 'center' }
      sheet.getColumn(rowNumber).font = { size: 14, family: 2 }
    });

    workbook.xlsx.writeBuffer().then(function (buffer) {
      saveAs(new Blob([buffer], {
        type: 'application/octet-stream'
      }), 'ABC' + '.' + 'xlsx');
    });

    // await workbook.xlsx.writeFile("D:/ipad项目/first.xlsx").then({

    // });

    // var FileSaver = require('file-saver');
    // var blob = new Blob(["Hello, world!"], { type: "text/plain;charset=utf-8" });
    // FileSaver.saveAs(blob, "hello world.txt");
  }

}
