export type ConversionMode = "smart" | "clean" | "pixel";
export type OcrLanguage = "eng+spa" | "eng" | "spa";

export interface ConversionSettings {
  mode: ConversionMode;
  language: OcrLanguage;
  preserveColor: boolean;
  preserveLayout: boolean;
}

export interface ProgressUpdate {
  value: number;
  phase: string;
  detail: string;
}

export interface ConversionResult {
  blob: Blob;
  fileName: string;
  pages: number;
  sourceKind: "pdf" | "image";
  usedOcr: boolean;
}

export interface PreviewData {
  url: string;
  kind: "pdf" | "image";
  pages?: number;
}
