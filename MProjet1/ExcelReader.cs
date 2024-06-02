using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using System.Data;
using Npgsql;

namespace MProjet1
{
    internal class ExcelReader
    {
        // The connection string should be to your PostgreSQL database.
        
        public void ReadExcelFile(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet articlesSheet = null;
            Excel.Worksheet achatsSheet = null;
            Excel.Worksheet ventesSheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(filePath);

                // Read "Articles" sheet
                articlesSheet = workbook.Sheets["Articles"];
                ReadWorksheet(articlesSheet);

                // Read "Achats" sheet
                achatsSheet = workbook.Sheets["Achats"];
                ReadWorksheet(achatsSheet);

                // Read "Ventes" sheet
                ventesSheet = workbook.Sheets["Ventes"];
                ReadWorksheet(ventesSheet);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                // Release other objects
                ReleaseObject(articlesSheet);
                ReleaseObject(achatsSheet);
                ReleaseObject(ventesSheet);
                ReleaseObject(workbooks);
            }
        }

        private void ReadWorksheet(Excel.Worksheet worksheet)
        {
            Excel.Range range = worksheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    object cellValue = (range.Cells[row, col] as Excel.Range).Value2;
                    // Process cellValue as needed
                }
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occurred while releasing object " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
        public void PrintExcelSheetsToConsole(string filePath)
        {
            var excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath);

            try
            {
                PrintWorksheet(workbook.Sheets["Articles"]);
                PrintWorksheet(workbook.Sheets["Achats"]);
                PrintWorksheet(workbook.Sheets["Ventes"]);
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        private void PrintWorksheet(Excel.Worksheet worksheet)
        {
            Console.WriteLine($"Sheet: {worksheet.Name}");
            Excel.Range range = worksheet.UsedRange;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    if (range.Cells[row, col] is Excel.Range cell)
                    {
                        Console.Write($"{cell.Value2}\t");
                    }
                }
                Console.WriteLine();
            }
            Console.WriteLine("--------------------------------------------------");
        }
        


        private void ReadWorksheetAndInsertData(Excel.Worksheet worksheet, string connectionString, string tableName)
        {
            Excel.Range range = worksheet.UsedRange;
            using (var conn = new NpgsqlConnection(connectionString))
            {
                conn.Open();
                string commandText;
                NpgsqlCommand cmd;

                for (int row = 2; row <= range.Rows.Count; row++) // Assuming the first row contains column headers
                {
                    if (tableName == "articles")
                    {
                        commandText = $"INSERT INTO \"{tableName}\" (id, libelle, pu) VALUES (@id, @libelle, @pu);";
                        cmd = new NpgsqlCommand(commandText, conn);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32((range.Cells[row, 1] as Excel.Range).Value2));
                        cmd.Parameters.AddWithValue("@libelle", (range.Cells[row, 2] as Excel.Range).Value2.ToString());
                        cmd.Parameters.AddWithValue("@pu", Convert.ToDecimal((range.Cells[row, 3] as Excel.Range).Value2));
                    }
                    else // "achats" or "ventes"
                    {
                        commandText = $"INSERT INTO \"{tableName}\" (num, id_art, qte) VALUES (@num, @id, @qte);";
                        cmd = new NpgsqlCommand(commandText, conn);
                        cmd.Parameters.AddWithValue("@num", Convert.ToInt32((range.Cells[row, 1] as Excel.Range).Value2));
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32((range.Cells[row, 2] as Excel.Range).Value2));
                        cmd.Parameters.AddWithValue("@qte", Convert.ToInt32((range.Cells[row, 3] as Excel.Range).Value2));
                        double excelDate = Convert.ToDouble((range.Cells[row, 4] as Excel.Range).Value2);
                        DateTime date = ConvertExcelDateToDateTime(excelDate);
                        cmd.Parameters.AddWithValue("@date", date);
                    }
                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void OpenExcelAndProcessSheets(string filePath, string connectionString, string tableName)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[tableName]; // Assuming the sheet name matches the table name

            try
            {
                ReadWorksheetAndInsertData(worksheet, connectionString, tableName);
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
        public static DateTime ConvertExcelDateToDateTime(double excelDate)
        {
            if (excelDate < 1)
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            // Excel's date system starts on January 1, 1900, but it is actually incremented by 2 due to a historical bug.
            return dateOfReference.AddDays(excelDate - 2);
        }

    }
}
