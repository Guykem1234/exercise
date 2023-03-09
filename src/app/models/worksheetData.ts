import { AddWorksheetOptions } from 'exceljs';

export interface WorkSheetData {
    name: string,
    data: any[],
    options?: Partial<AddWorksheetOptions>
}