import {
  Injectable,
  InternalServerErrorException,
  UnprocessableEntityException,
} from '@nestjs/common';
import { Cell, Row, Workbook } from 'exceljs';

@Injectable()
export class AppService {
  async readExcel(file: string) {
    try {
      const workbook = new Workbook();
      const readWorkbook = await workbook.xlsx.readFile(`assets/${file}`);
      const worksheet = readWorkbook.getWorksheet(1);
      const worksheetName = worksheet.name;
      const firstRow = worksheet.getRow(1);
      let data = [];

      if (!firstRow.cellCount) return;

      worksheet.eachRow(
        { includeEmpty: true },
        (row: Row, rowIndex: number) => {
          if (rowIndex === 1) return;

          const values = row.values;
          const object = {};

          for (let index = 1; index < firstRow.values.length; index++) {
            object[firstRow.values[index]] = values[index] ?? '';
          }

          data = [...data, object];
        },
      );

      const headers = Object.keys(data[0]);

      return { worksheetName, headers, data };
    } catch (error) {
      if (error?.code === 'EBUSY') {
        throw new UnprocessableEntityException(
          'File is currently open or locked!',
        );
      }
      throw new InternalServerErrorException('Error: Unable to process!');
    }
  }

  async overwriteExcel(file: string, entries: any[]) {
    try {
      const workbook = new Workbook();
      const readWorkbook = await workbook.xlsx.readFile(`assets/${file}`);
      const worksheet = readWorkbook.getWorksheet(1);
      const columnCount = worksheet.columnCount;
      const styles = [];

      worksheet.getRow(2).eachCell({ includeEmpty: true }, (cell: Cell) => {
        styles.push(cell.style);
      });

      worksheet.eachRow((row: Row, rowIndex) => {
        if (rowIndex === 1) return;
        row.splice(1, columnCount);
        row.commit();
      });

      let counter = 2;
      for (const entry of entries) {
        const normalizeEntry = Object.values(entry).map((value: any) => {
          if (Date.parse(value) && typeof value === 'string') {
            return new Date(value);
          }
          return value;
        });

        worksheet.getRow(counter).values = normalizeEntry;
        worksheet.getRow(counter).eachCell((cell: Cell, cellIndex: number) => {
          if (styles[cellIndex - 1]?.numFmt) {
            cell.numFmt = styles[cellIndex - 1].numFmt;
          }

          if (styles[cellIndex - 1]?.font) {
            cell.font = styles[cellIndex - 1].font;
          }

          if (styles[cellIndex - 1]?.border) {
            cell.border = styles[cellIndex - 1].border;
          }

          if (styles[cellIndex - 1]?.alignment) {
            cell.alignment = styles[cellIndex - 1].alignment;
          }

          if (styles[cellIndex - 1]?.fill) {
            cell.fill = styles[cellIndex - 1].fill;
          }
        });

        counter++;
      }

      const options = { useSharedStrings: true, useStyles: true };
      await readWorkbook.xlsx.writeFile(`assets/${file}`, options);
      return { message: 'File is successfully saved.' };
    } catch (error) {
      if (error?.code === 'EBUSY') {
        throw new UnprocessableEntityException(
          'File is currently open or locked!',
        );
      }
      throw new InternalServerErrorException('Error: Unable to process!');
    }
  }
}
