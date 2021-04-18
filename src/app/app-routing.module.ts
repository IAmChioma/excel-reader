import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ExcelReaderComponent } from './excel-reader/excel-reader.component';

const routes: Routes = [
  {path: '', component: ExcelReaderComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
