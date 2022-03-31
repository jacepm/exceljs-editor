import {
  Injectable,
  InternalServerErrorException,
  UnprocessableEntityException,
} from '@nestjs/common';
import { Row, Workbook } from 'exceljs';

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
}
