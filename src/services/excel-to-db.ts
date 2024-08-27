import * as xlsx from "xlsx";
import { DatabaseFactory, DatabaseConfig } from "../database/database-factory";
import { Database } from "../database/database";

export class ExcelToDB {
  private db: Database;

  constructor(dbConfig: DatabaseConfig) {
    this.db = DatabaseFactory.createDatabase(dbConfig);
  }

  public async uploadExcel(
    filePath: string,
    table: string,
    mapping: string[] = []
  ): Promise<void> {
    try {
      const data: string[] = [];
      const keys: string[] = [];
      const mappingWasProvided = mapping.length > 0;
      let isHeaderProcessed = false;

      await this.db.connect();

      const workbook: xlsx.WorkBook = xlsx.readFile(filePath);
      const sheetName: string = workbook.SheetNames[0];
      const sheet: xlsx.WorkSheet = workbook.Sheets[sheetName];
      const rows: Record<string, string>[] = xlsx.utils.sheet_to_json(sheet, {
        raw: true,
      });

      rows.forEach((row) => {
        const entries = Object.entries(row);
        const rowData: string[] = [];

        entries.forEach(([key, value]) => {
          if (mappingWasProvided && !mapping.includes(key)) return;
          if (!isHeaderProcessed) keys.push(key);
          rowData.push(`'${value}'`);
        });

        if (rowData.length > 0) {
          data.push(`(${rowData.join(", ")})`);
        }

        isHeaderProcessed = true;
      });

      const queryData = data.join(", ");
      const queryKeys = keys.join(", ");

      await this.db.insertData(table, queryData, queryKeys);
    } catch (error) {
      throw error;
    } finally {
      await this.db.closeConnection();
    }
  }
}
