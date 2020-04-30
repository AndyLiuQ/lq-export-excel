import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { CsvModule } from '@ctrl/ngx-csv';
import { ExportExcelUtil } from 'src/common/utils/export-excel-util';


import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    CsvModule
  ],
  providers: [ExportExcelUtil],
  bootstrap: [AppComponent]
})
export class AppModule { }
