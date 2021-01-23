using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

/* Excel Reading Class 
 * Tested against memory leaks
 * Opens and removes all traces of microsoft excels
 * Does not lock out the user from opening the excel spreadsheet
 * Sample/Intended Example:
 * 
 *          string excelfile = ""; // The directory for the excel doc
 *          string sheetname = ""; // The name of the sheet
 *          
 *          Excel xlapp = new Excel();
 *          try
 *          {
 *              xlapp.Open(excelfile, sheetname);
 *              xlapp.ReadCell01(1, 1);
 *          }
 *          finally
 *          {
 *              xlapp.Closed();
 *              GC.Collect();
 *              GC.WaitForPendingFinalizers()
 *              GC.Collect();
 *          }
*/
namespace Goat_Therapy_WPI_Hackathon
{
   class Excel
    {
        public _Application excelApp;
        public _Excel.Workbook wb;
        public _Excel.Worksheet ws;

        public void Open(string pth, string sheet) // Open the spreadsheet
        {
            excelApp = new _Excel.Application();
            var wbs = excelApp.Workbooks;
            wb = wbs.Open(pth);
            ws = wb.Worksheets[sheet];

        }
        public void Close_Open (string sheet) // Close current sheet and open the input
        {
            while (Marshal.ReleaseComObject(ws) != 0) { }
            ws = null;
            ws = wb.Worksheets[sheet];
        }

        public void Closed() // Close all traces of the spreadsheet
        {
            wb.Close(false);
            excelApp.Quit();


            //Manual disposal because of COM
            while (Marshal.ReleaseComObject(excelApp) != 0) { }
            while (Marshal.ReleaseComObject(wb) != 0) { }
            while (Marshal.ReleaseComObject(ws) != 0) { }
            //Manual disposal because of COM

            excelApp = null;
            wb = null;
            ws = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
       
        public double ReadCell (int j, int i) // Read Number
        {
            
            
            if(ws.Cells[i, j].Value2 != null)
            {

                return ws.Cells[i, j].Value2; 
            }
           
            else
            {
                return 0;
            }
            
        }
        public string ReadCell01(int j, int i) // Read String
        {


            if (ws.Cells[i, j].Value2 != null)
            {

                return ws.Cells[i, j].Value2;
            }

            else
            {
                return "";
            }

        }

      

    }
}
