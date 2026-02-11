export interface SignatureFile {
  name: string; // The base name, e.g., "HongGilDong"
  variant: string; // The full filename or id, e.g., "HongGilDong_1"
  previewUrl: string; // Changed from dataUrl: Use Blob URL for memory efficiency
  width: number;
  height: number;
}

export interface CellData {
  value: string | number | null;
  address: string; // e.g., "A1"
  row: number;
  col: number;
}

export interface RowData {
  index: number; // 1-based row index
  cells: CellData[];
}

export interface SheetData {
  name: string;
  rows: RowData[];
  mergedCells?: string[]; // Array of merged cell ranges like "A1:B2"
  printArea?: string; // Print area range like "A1:Z100"
}

export interface SignatureAssignment {
  row: number;
  col: number;
  signatureBaseName: string; // "HongGilDong"
  signatureVariantId: string; // Specific file ID used
  rotation: number; // Degrees -5 to 5
  scale: number; // Percentage 0.95 to 1.05
  offsetX: number; // Pixels
  offsetY: number; // Pixels
}

export interface AppState {
  step: 'upload' | 'preview' | 'export';
  excelFile: File | null;
  excelBuffer: ArrayBuffer | null;
  sheetData: SheetData | null;
  signatures: Map<string, SignatureFile[]>; // Map "HongGilDong" -> [File1, File2]
  assignments: Map<string, SignatureAssignment>; // Map "Row:Col" -> Assignment
}