declare var require: any
var _ = require('lodash');
import { Directive } from '@angular/core';
import { Workbook } from 'exceljs';
import * as filesaver from 'file-saver';
import { DatePipe } from '@angular/common';
import * as exceldata from './test-data/test1.json';
import {XlsxData} from './interface/xlsx-data';
const CSV_TYPE = 'application/vnd.ms-excel';
const CSV_EXTENSION = '.xlsx';

const defaults = {
  "worksheetName": null,
  "image": null,
  "title": null,
  "tables": null
}
@Directive({
  selector: '[appXlsxWriter]'
})
export class XlsxWriterDirective {

  constructor(private datePipe: DatePipe) { }
  ngOnInit(){
    this.generateExcel();
  }
  public generateExcel() {
   
    let data  = <XlsxData> exceldata;
    let xlxdata = {...defaults, ...data['default']};
    console.log('theData', xlxdata);
    const workbook = new Workbook();
    //worksheetName
    const worksheetName = this.getWorksheetName(xlxdata.worksheetName);
    const worksheet = workbook.addWorksheet(worksheetName);
    
    if (xlxdata.image) {
      _.forEach(xlxdata.image.data, image=>{
        if (this.isImagePresent(image)) {
          this.styleImage(workbook, worksheet, image);
        }
      })
    }

    if (xlxdata.title) {
      _.forEach(xlxdata.title.data, title=>{
         title = xlxdata.title.data[0];
        if (this.isTitlePresent(title)) {
          this.cellMergeAndStyle(worksheet, title);
        }
      })
    }
    //tables  
    if (xlxdata.tables) {
      const tables = xlxdata.tables;

      tables.forEach(table => {
        table.headers.data.forEach(header => {
          this.cellMergeAndStyle(worksheet, header);
        });

        table.rowsData.forEach(row => {
          row.forEach(cell => {
            this.cellMergeAndStyle(worksheet, cell);
          })
        });
        worksheet.addRow([]);
      });
    }

    //column width
    worksheet.columns.forEach(column => {
      column.width = 20;
    });

    //save
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], { type: CSV_TYPE });
      filesaver.saveAs(blob, worksheetName + CSV_EXTENSION);
    });

  }
  getWorksheetName(name: string): string {
    return name ? name : "Report";
  }
  isImagePresent(data): boolean {
    return data.name ? true : false;
  }
  isTitlePresent(data): boolean {
    return data.name ? true : false;
  }

  styleImage(workbook, worksheet, image) {
    const logo = workbook.addImage({
      base64: image.name,
      extension: 'jpeg',
    });

    const topLeft = image['topLeft'];
    const bottomRight = image['bottomRight'];
    worksheet.addImage(logo, {
      tl: { col: topLeft.col, row: topLeft.row },
      br: { col: bottomRight.col, row: bottomRight.row }
    });
  }

  cellMergeAndStyle(worksheet, data) {
    let keys = Object.keys(data);
    if (keys.length) {
      if (_.includes(keys, "mergeCells")) {
        let mergeCellsKeys = Object.keys(data['mergeCells']);

        if (_.includes(mergeCellsKeys, "start")) {
          const start = data['mergeCells'].start;
          const end = _.includes(mergeCellsKeys, "end") ? data['mergeCells'].end : start;
          worksheet.mergeCells(start, end);

          if (_.includes(keys, "name"))
            worksheet.getCell(start).value = data['name'];
          worksheet.getCell(start).alignment = { horizontal: 'center' };

          if (_.includes(keys, "style")) {
            if (Object.keys(data['style']).length) {
              let cellProperties = {};
              let fontProperties = {};
              _.forEach(data['style'], (value, key) => {
                if (key === "border") {
                  if (value === "true") {
                    const a = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
                    worksheet.getCell(start).border = a;
                  }
                }
                else if (key === "bgColor" || key === "fgColor") {
                  let color = {
                    'argb': value
                  };
                  cellProperties[key] = color;
                }
                else {
                  if (key === "color") {
                    let color = {
                      'argb': value
                    };
                    fontProperties[key] = color;
                  }
                  else fontProperties[key] = value;
                }
              });


              if (Object.keys(fontProperties).length)
                worksheet.getCell(start).font = fontProperties;

              if (Object.keys(cellProperties).length) {
                cellProperties['type'] = "pattern";
                cellProperties['pattern'] = "solid";
                worksheet.getCell(start).fill = cellProperties;
              }
            }
          }
        }
      }
    }
  }

}
