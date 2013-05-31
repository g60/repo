using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Car_Price_Guider
{

    class AuctionCatalogueParser_Newark
    {

        private string _project_code;
        private string _filePath;
        private Excel.Worksheet _catalogueWorksheet;
        private Excel.Application ExcelObj;
        private Excel.Workbook theWorkbook;

        public AuctionCatalogueParser_Newark(string FileName)
        {
            _filePath = FileName;
        }

        /// <summary>
        /// Opens the spreadsheet ready for reading
        /// </summary>
        public void OpenSpreadsheet()
        {
            try
            {
                ExcelObj = new Excel.Application();

                theWorkbook = ExcelObj.Workbooks.Open(_filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                
                Excel.Sheets sheets = theWorkbook.Worksheets;

                _catalogueWorksheet = (Excel.Worksheet)sheets.get_Item(1);
                
            }
            catch
            {
                CloseSpreadsheet();
                throw;
            } // end try / catch
        }

        /// <summary>
        /// Closes the currently opened spreadsheet
        /// </summary>
        public void CloseSpreadsheet()
        {
            try
            {
                // Repeat xl.Visible and xl.UserControl releases just to be sure
                // we didn't error out ahead of time.

                if (ExcelObj != null)
                {
                    ExcelObj.Visible = false;
                    ExcelObj.UserControl = false;
                } // end if

                if (theWorkbook != null)
                {
                    // Close the document and avoid user prompts to save if our method failed.
                    theWorkbook.Close(false, null, null);
                    ExcelObj.Workbooks.Close();
                } // end if
            }
            catch { }

            // Gracefully exit out and destroy all COM objects to avoid hanging instances
            // of Excel.exe whether our method failed or not.

            if (theWorkbook != null) { Marshal.ReleaseComObject(theWorkbook); }
            if (ExcelObj != null) { ExcelObj.Quit(); }
            if (ExcelObj != null) { Marshal.ReleaseComObject(ExcelObj); }

            theWorkbook = null;
            ExcelObj = null;
            GC.Collect();
        }

        public static List<CarDetails> ParseCatalogue_StoredFile(string FileName)
        {
            Excel.Application ExcelObj = null;

            Excel.Workbook theWorkbook = null;

            List<CarDetails> returnDetails = new List<CarDetails>();

            try {

                ExcelObj = new Excel.Application();

                theWorkbook = ExcelObj.Workbooks.Open(FileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                
                Excel.Sheets sheets = theWorkbook.Worksheets;

                Excel.Worksheet catalogueWorksheet = (Excel.Worksheet)sheets.get_Item(1);

            

                const int C_COL_NUM__MAKE = 4; // D
                const int C_COL_NUM__MODEL = 6; // F
                const int C_COL_NUM__TYPE = 7; // G
                const int C_COL_NUM__REGISTERED = 8; // H
                const int C_COL_NUM__FUEL = 10; // J
                const int C_COL_NUM__TRANS = 11; // K
                const int C_COL_NUM__DOORS = 13; // M
                const int C_COL_NUM__MILES = 14; // N


                Excel.Range range = catalogueWorksheet.get_Range("A1", "R50");

                System.Array myvalues = (System.Array)range.Cells.Value;


                // loop through each depot row
                for (int rowIndex = 1; rowIndex <= range.Cells.Rows.Count; rowIndex++)
                {
                    Console.WriteLine("**********************************");
                    Console.WriteLine("MAKE:" + (myvalues.GetValue(rowIndex, C_COL_NUM__MAKE) ?? "").ToString().Trim());
                    Console.WriteLine("MODEL:" + (myvalues.GetValue(rowIndex, C_COL_NUM__MODEL) ?? "").ToString().Trim());
                    Console.WriteLine("TYPE:" + (myvalues.GetValue(rowIndex, C_COL_NUM__TYPE) ?? "").ToString().Trim());
                    Console.WriteLine("REGDATE:" + (myvalues.GetValue(rowIndex, C_COL_NUM__REGISTERED) ?? "").ToString().Trim());
                    Console.WriteLine("FUEL:" + (myvalues.GetValue(rowIndex, C_COL_NUM__FUEL) ?? "").ToString().Trim());
                    Console.WriteLine("TRANS:" + (myvalues.GetValue(rowIndex, C_COL_NUM__TRANS) ?? "").ToString().Trim());
                    Console.WriteLine("DOORS:" + (myvalues.GetValue(rowIndex, C_COL_NUM__DOORS) ?? "").ToString().Trim());
                    Console.WriteLine("MILES:" + (myvalues.GetValue(rowIndex, C_COL_NUM__MILES) ?? "").ToString().Trim());
                    Console.WriteLine("**********************************");

                    CarDetails newCar = new CarDetails();
                    newCar.Make = (myvalues.GetValue(rowIndex, C_COL_NUM__MAKE) ?? "").ToString().Trim();


                } // end for

                if (ExcelObj != null)
                {
                    ExcelObj.Visible = false;
                    ExcelObj.UserControl = false;
                } // end if

                if (theWorkbook != null)
                {
                    // Close the document and avoid user prompts to save if our method failed.
                    theWorkbook.Close(false, null, null);
                    ExcelObj.Workbooks.Close();
                } // end if

            }
            finally
            {
                if (theWorkbook != null) { Marshal.ReleaseComObject(theWorkbook); }
                if (ExcelObj != null) { ExcelObj.Quit(); }
                if (ExcelObj != null) { Marshal.ReleaseComObject(ExcelObj); }

                theWorkbook = null;
                ExcelObj = null;
                GC.Collect();
            }

            return returnDetails;

        }

    }
}
