export type SheetDto = {
  name: string;
  rows: string[][];
};

export type WorkbookDto = {
  fileName: string;
  sheets: SheetDto[];
};
