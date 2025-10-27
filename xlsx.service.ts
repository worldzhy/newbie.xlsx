import {Injectable} from '@nestjs/common';
import {Express} from 'express';
import * as XLSX from 'xlsx';

@Injectable()
export class XLSXService {
  private workbook: XLSX.WorkBook;

  createWorkbook() {
    this.workbook = XLSX.utils.book_new();
  }

  getSheets(): string[] {
    return this.workbook.SheetNames;
  }

  //**********************/
  //* Load file          */
  //**********************/

  loadLocalFile(filePath: string) {
    this.workbook = XLSX.readFile(filePath);
  }

  loadFile(file: Express.Multer.File) {
    this.workbook = XLSX.read(file.buffer);
  }

  loadBuffer(buffer: Buffer) {
    this.workbook = XLSX.read(buffer);
  }

  //************************/
  //* Get columns and rows */
  //************************/

  getColumnsBySheetIndex(index: number): string[] {
    const sheet = this.workbook.Sheets[this.workbook.SheetNames[index]];

    return this.getColumns(sheet);
  }

  getColumnsBySheetName(name: string): string[] {
    const sheet = this.workbook.Sheets[name];

    return this.getColumns(sheet);
  }

  getRowsBySheetIndex(index: number): object[] {
    const sheet = this.workbook.Sheets[this.workbook.SheetNames[index]];
    return this.getRows(sheet);
  }

  getRowsBySheetName(name: string): object[] {
    const sheet = this.workbook.Sheets[name];
    return this.getRows(sheet);
  }

  //**********************/
  //* Save file          */
  //**********************/

  writeFile(fileName: string) {
    XLSX.writeFile(this.workbook, fileName);
  }

  //**********************/
  //* Private operations */
  //**********************/

  private getColumns(sheet: XLSX.WorkSheet): string[] {
    const columns: string[] = [];

    if (sheet['!ref']) {
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const startRow = range.s.r; // start in the first row

      // walk every column in the range
      for (let C = range.s.c; C <= range.e.c; ++C) {
        let column = 'UNKNOWN ' + C; // <-- replace with your desired default

        const cell = sheet[XLSX.utils.encode_cell({c: C, r: startRow})]; // find the cell in the first row
        if (cell && cell.t) {
          column = XLSX.utils.format_cell(cell);
        }

        columns.push(column);
      }
    }

    return columns;
  }

  private getRows(sheet: XLSX.WorkSheet): object[] {
    return XLSX.utils.sheet_to_json(sheet);
  }
}
