import { ExcelToDB } from "./index";

async function main() {
  // Database configuration
  const dbConfig = {
    type: "postgres" as "postgres" | "sqlserver", // or "sqlserver"
    config: {
      host: "localhost",
      user: "sa",
      password: "password123",
      // database: "postgres",
      port: 5432,
    },
  };

  // Mapping between Excel columns and database fields
  const mapping = ["id", "name", "phone", "email"];

  // Create an instance of ExcelToDB with the database configuration

  const excelToDB = new ExcelToDB(dbConfig);

  try {
    // Upload the Excel data to the database
    console.log("Start file processing ");
    await excelToDB.uploadExcel("./.vscode/Book1.xlsx", "Employee", mapping);
    console.log("Upload complete");
  } catch (error) {
    console.error("Error:", error);
  }
}

// Run the main function to test the code
main().catch(console.error);
