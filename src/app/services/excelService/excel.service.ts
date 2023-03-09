import { Injectable } from '@angular/core';


import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';
import { WorkSheetData } from 'src/app/models/worksheetData';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() { }

  createWorkBook(...worksheets: WorkSheetData[]) {
    const workbook = new Workbook();
    for (let i = 0; i < worksheets.length; i++) {
      const worksheetData = worksheets[i];
      console.log(worksheetData.options)
      const ws = workbook.addWorksheet(worksheetData.name, worksheetData.options)

      ws.columns = Object.keys(worksheetData.data[0]).map(key => ({ header: key, key: key, width: 50 }))
      ws.addRows(worksheetData.data)
      ws.columns.forEach(col => {
        col.width = 50;
        col.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
      })


      ws.columns[1].eachCell((cell, rowNumber) => {
        console.log(rowNumber)
        if (rowNumber == 1) {
          cell.fill = {
            type: 'pattern',
            pattern: 'darkTrellis',
            fgColor: { argb: 'CCFFC0FF' },
            bgColor: { argb: 'CCFFC0FF' }
          }
          return
        }
        cell.fill = {
          type: 'pattern',
          pattern: 'darkTrellis',
          fgColor: { argb: 'FFFFFF00' },
          bgColor: { argb: 'FF0000FF' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      })

    }

    workbook.xlsx.writeBuffer().then(function (buffer) {
      //send file to user
      const data: Blob = new Blob([buffer], {
        type: EXCEL_TYPE
      });
      saveAs(data, "name" + EXCEL_EXTENSION);
    });
  }

}
