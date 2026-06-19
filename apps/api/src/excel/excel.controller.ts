import { BadRequestException, Body, Controller, Get, Header, Post, Res, UploadedFile, UseInterceptors } from "@nestjs/common";
import { FileInterceptor } from "@nestjs/platform-express";
import type { Response } from "express";
import { ExcelService } from "./excel.service";
import type { WorkbookDto } from "./dto/workbook.dto";

@Controller("excel")
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Get("health")
  health() {
    return { status: "ok" };
  }

  @Post("parse")
  @UseInterceptors(FileInterceptor("file"))
  async parse(@UploadedFile() file?: Express.Multer.File) {
    if (!file) {
      throw new BadRequestException("Debes enviar un archivo en el campo 'file'.");
    }

    return this.excelService.parse(file.originalname, file.buffer);
  }

  @Post("export")
  @Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
  async export(@Body() workbook: WorkbookDto, @Res() response: Response) {
    const buffer = await this.excelService.export(workbook);
    const fileName = this.excelService.exportFileName(workbook.fileName);

    response.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(fileName)}"`);
    response.send(buffer);
  }
}
