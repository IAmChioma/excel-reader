import { Component, OnInit } from '@angular/core';
import * as wjcCore from '@grapecity/wijmo';
import * as wjcXlsx from '@grapecity/wijmo.xlsx';
@Component({
  selector: 'app-excel-reader',
  templateUrl: './excel-reader.component.html',
  styleUrls: ['./excel-reader.component.scss']
})
export class ExcelReaderComponent {


  workbook: wjcXlsx.Workbook;
  sheetIndex: number;
  //
  ngAfterViewInit() {
    document.getElementById('importFile').addEventListener('change', () => {
      this._loadWorkbook();
    });
  }
  //
  tabClicked(e: MouseEvent, index: number) {
    e.preventDefault();
    this._drawSheet(index);
  }
  //
  private _loadWorkbook() {
    let reader = new FileReader();
    //
    reader.onload = (e) => {
      let workbook = new wjcXlsx.Workbook();
      workbook.loadAsync(<string>reader.result, (result: wjcXlsx.Workbook) => {
        this.workbook = result;
        this._drawSheet(this.workbook.activeWorksheet || 0);
      });
    };
    //
    let file = (<HTMLInputElement>document.getElementById('importFile')).files[0];
    if (file) {
      reader.readAsDataURL(file);
    }
  }
  //
  private _drawSheet(sheetIndex: number) {
    let drawRoot = document.getElementById('tableHost');
    drawRoot.textContent = '';
    this.sheetIndex = sheetIndex;
    this._drawWorksheet(this.workbook, sheetIndex, drawRoot, 200, 100);
  }
  //
  private _drawWorksheet(workbook: wjcXlsx.IWorkbook, sheetIndex: number, rootElement: HTMLElement, maxRows: number, maxColumns: number) {
    if (!workbook || !workbook.sheets || sheetIndex < 0 || workbook.sheets.length == 0) {
      return;
    }
    //
    sheetIndex = Math.min(sheetIndex, workbook.sheets.length - 1);
    //
    if (maxRows == null) {
      maxRows = 200;
    }
    if (maxColumns == null) {
      maxColumns = 100;
    }
    
    // Namespace and XlsxConverter shortcuts.
    let sheet = workbook.sheets[sheetIndex],
      defaultRowHeight = 20,
      defaultColumnWidth = 60,
      tableEl = document.createElement('table');
    //
    tableEl.border = '1';
    tableEl.style.borderCollapse = 'collapse';
    tableEl.style.display = 'table-striped';
    //
    let maxRowCells = 0;
    for (let r = 0; sheet.rows && r < sheet.rows.length; r++) {
      if (sheet.rows[r] && sheet.rows[r].cells) {
        maxRowCells = Math.max(maxRowCells, sheet.rows[r].cells.length);
      }
    }
    //
    // add columns
    let columnCount = 0;
    if (sheet.columns) {
      columnCount = sheet.columns.length;
      maxRowCells = Math.min(Math.max(maxRowCells, columnCount), maxColumns);
      for (let c = 0; c < maxRowCells; c++) {
        let col = c < columnCount ? sheet.columns[c] : null,
          colEl = document.createElement('col');
        tableEl.appendChild(colEl);
        let colWidth = defaultColumnWidth + 'px';
        if (col) {
          if (col.autoWidth) {
            colWidth = '';
          } else if (col.width != null) {
            colWidth = col.width + 'px';
          }
        }
        colEl.style.width = colWidth;
      }
    }
    //
    // add rows
    let rowCount = Math.min(maxRows, sheet.rows.length);
    for (let r = 0; sheet.rows && r < rowCount; r++) {
      let row = sheet.rows[r],
        rowEl = document.createElement('tr');
      tableEl.appendChild(rowEl);
      if (row) {
        if (row.height != null) {
          rowEl.style.height = row.height + 'px';
        }
        for (let c = 0; row.cells && c < row.cells.length; c++) {
          let cell = row.cells[c],
            cellEl = document.createElement('td');
          rowEl.appendChild(cellEl);
          if (cell) {
            let value = cell.value;
            if (!(value == null || value !== value)) { // TBD: check for NaN should be eliminated
              if (wjcCore.isString(value) && value.charAt(0) == "'") {
                value = value.substr(1);
              }
              let netFormat = '';
              if (cell.style && cell.style.format) {
                netFormat = wjcXlsx.Workbook.fromXlsxFormat(cell.style.format)[0];
              }
              let fmtValue = netFormat ? wjcCore.Globalize.format(value, netFormat) : value;
              cellEl.innerHTML = wjcCore.escapeHtml(fmtValue);
            }
            if (cell.colSpan && cell.colSpan > 1) {
              cellEl.colSpan = cell.colSpan;
              c += cellEl.colSpan - 1;
            }
          }
        }
      }
      // pad with empty cells
      let padCellsCount = maxRowCells - (row && row.cells ? row.cells.length : 0);
      for (let i = 0; i < padCellsCount; i++) {
        rowEl.appendChild(document.createElement('td'));
      }
      //
      if (!rowEl.style.height) {
        rowEl.style.height = defaultRowHeight + 'px';
      }
    }
    //
    // append child to table
    rootElement.appendChild(tableEl);
  }
  //
}

