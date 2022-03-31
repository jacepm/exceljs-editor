import { Body, Controller, Get, Post, Query } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get('excel')
  async readExcel(@Query('file') file: string) {
    return await this.appService.readExcel(file);
  }
}
