import { Component } from '@angular/core';
import { ngCsv } from 'ng-csv';
import * as XLSX from 'xlsx';
import { ExportExcelUtil } from 'src/common/utils/export-excel-util';
import { ExceljsUtilService } from 'src/common/utils/exceljs-util.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})

/**
 * @author AndyLiu
 */
export class AppComponent {
  

  constructor() { 
  }

  testExcelJS(){
    let excejs = new ExceljsUtilService();
    excejs.test();
  }

}

