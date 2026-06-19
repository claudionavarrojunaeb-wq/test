import { BadRequestException, Injectable } from "@nestjs/common";
import ExcelJS from "exceljs";
import { Readable } from "node:stream";
import type { SheetDto, WorkbookDto } from "./dto/workbook.dto";

@Injectable()
export class ExcelService {
  async parse(fileName: string, buffer: Buffer): Promise<WorkbookDto> {
    try {
      const workbook = new ExcelJS.Workbook();

      if (fileName.toLowerCase().endsWith(".csv")) {
        await workbook.csv.read(Readable.from([buffer]));
      } else {
        await workbook.xlsx.load(buffer as unknown as Parameters<typeof workbook.xlsx.load>[0]);
      }

      const sheets = workbook.worksheets.map((worksheet) => this.sheetToDto(worksheet));

      if (!sheets.length) {
        throw new BadRequestException("El archivo no contiene hojas.");
      }

      return { fileName, sheets };
    } catch (error) {
      if (error instanceof BadRequestException) throw error;
      throw new BadRequestException("No se pudo leer el archivo XLSX o CSV.");
    }
  }

  async export(workbook: WorkbookDto): Promise<Buffer> {
    if (!workbook?.sheets?.length) {
      throw new BadRequestException("Debes enviar al menos una hoja para exportar.");
    }

    const output = new ExcelJS.Workbook();

    const usedNames = new Set<string>();

    for (const sheet of workbook.sheets) {
      if (!sheet.name) {
        throw new BadRequestException("Todas las hojas deben tener nombre.");
      }

      const worksheet = output.addWorksheet(this.safeSheetName(sheet.name, usedNames));
      for (const row of this.trimTrailingEmpty(sheet.rows ?? [])) {
        worksheet.addRow(row);
      }
    }

    const data = await output.xlsx.writeBuffer();
    return Buffer.from(data);
  }

  exportFileName(fileName: string) {
    const baseName = (fileName || "libro")
      .replace(/\.[^.]+$/, "")
      .replace(/[\r\n"\\/:*?<>|]+/g, " ")
      .trim();
    return `${baseName || "libro"}-editado.xlsx`;
  }

  private safeSheetName(name: string, usedNames: Set<string>): string {
    const baseName = name.replace(/[\\/?*:[\]]/g, " ").trim() || "Hoja";
    let safeName = baseName.slice(0, 31);
    let suffix = 1;

    while (usedNames.has(safeName)) {
      const counter = ` ${suffix}`;
      safeName = `${baseName.slice(0, 31 - counter.length)}${counter}`;
      suffix += 1;
    }

    usedNames.add(safeName);
    return safeName;
  }

  private sheetToDto(worksheet: ExcelJS.Worksheet): SheetDto {
    const rows: string[][] = [];
    const columnCount = Math.max(worksheet.columnCount, 1);

    for (let rowIndex = 1; rowIndex <= Math.max(worksheet.rowCount, 1); rowIndex += 1) {
      const row = worksheet.getRow(rowIndex);
      rows.push(
        Array.from({ length: columnCount }, (_, columnIndex) => {
          const cell = row.getCell(columnIndex + 1);
          return String(cell.text || cell.value || "");
        }),
      );
    }

    return {
      name: worksheet.name,
      rows: this.normalizeRows(rows),
    };
  }

  private normalizeRows(rows: string[][]): string[][] {
    const normalized = rows.length ? rows.map((row) => row.map((value) => String(value ?? ""))) : [[""]];
    const width = Math.max(...normalized.map((row) => row.length), 1);

    return normalized.map((row) => {
      while (row.length < width) row.push("");
      return row;
    });
  }

  private trimTrailingEmpty(rows: string[][]): string[][] {
    return rows.map((row) => {
      let lastValue = row.length - 1;
      while (lastValue >= 0 && row[lastValue] === "") lastValue -= 1;
      return row.slice(0, lastValue + 1);
    });
  }
}
