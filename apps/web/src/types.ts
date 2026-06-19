export type Sheet = {
  name: string;
  rows: string[][];
};

export type WorkbookPayload = {
  fileName: string;
  sheets: Sheet[];
};
