import { Component } from '@angular/core';
import { ExcelService } from './services/excelService/excel.service';
import { WorkSheetData } from './models/worksheetData';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {



  constructor(private excelService: ExcelService) {
    console.log(new Date(1995, 11, 17) )
    const data = [{
      caseWorked: "abc",
      note: "Test fsdfsdf sdfsdfsdfsdfsdfasd fasdfasdfasdfasdf dsfasdfasdfsadf asdfasdfasdfdfsdfasdf asdfasdfasdfasdf sdf",
      id: "1234",
      date: new Date(1995, 11, 17) 
    },
    {
      caseWorked: "def",
      note: "test 1",
      id: "1234",
      date: new Date(1996, 0, 17) 
    },
    {
      caseWorked: "def",
      note: "Test 2",
      id: "3456",
      date: new Date(1996, 10, 17) 
    }];
    const workSheet: WorkSheetData = {
      name: "first",
      data,
      options: { pageSetup: { fitToPage: true, fitToHeight: 5, fitToWidth: 7, horizontalCentered: true, verticalCentered: true } }
    }
    excelService.createWorkBook(workSheet)
  }




  title = 'ExeclJs';
}
