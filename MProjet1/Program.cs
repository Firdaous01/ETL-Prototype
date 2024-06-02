using ETLProcess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MProjet1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            Application.EnableVisualStyles();
            

            // Create an instance of the ExcelReader class
            ExcelReader reader = new ExcelReader();

            // Specify the path to the Excel file
            string filePath = @"C:\Users\Lenovo\Documents\LST AD\S6\ing des donnees\MProjet1\MProjet1\registre.xls";

            // Call the method to print Excel sheets to the console
            reader.PrintExcelSheetsToConsole(filePath);

            string connectionString = "Host=localhost;Username=postgres;Password=root;Database=Registre";

            string tableName = "articles";
            string tableName2 = "achats";
            string tableName3 = "ventes";
            reader.OpenExcelAndProcessSheets(filePath, connectionString, tableName);
            reader.OpenExcelAndProcessSheets(filePath, connectionString, tableName2);
            reader.OpenExcelAndProcessSheets(filePath, connectionString, tableName3);
            //Transformation etl = new Transformation(connectionString);
            //etl.ExtractTransformLoad();
            // Create an instance of your ETL class
            
            DataETL etlProcessor = new DataETL();

            // Call the RunETL method to execute your ETL process
            etlProcessor.RunETL();

            Console.WriteLine("ETL process completed successfully.");
            Console.ReadLine(); // Keep the console open to view the results

        }
    }
}
