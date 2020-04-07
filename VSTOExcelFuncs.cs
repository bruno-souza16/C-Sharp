using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFuncs
{
    class MngtExcel
    { 
        // Instantiate Excel App
        public _Application InstanceExcel()
        {
            _Application excel = new _Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            return excel;
        }

        // Open a Excel file and stores in a Workbook variable
        // typeFile -> 0 to Excel, 1 to CSV, txt ---- filename --> fullpath ----- excel --> instance of Excel App
        public Workbook OpenFile(int typeFile, string filename, _Application excel) 
        {
            if(typeFile == 0)
            {
               Workbook w = excel.Workbooks.Open(filename);
            }
            else if(Type == 1)
            {
                Workbook w = excel.Workbooks.OpenXML(filename);
            }
            else
            {
                Workbook w = null;
            }
            return w;
        }

        // Get a worksheet of workbook selected and return to stores in a Worksheet variable.
        // workbook --> Excel file instantiated, ws_order --> number of sheet in Excel file.
        public Worksheet SelectSheet(Workbook workbook, int ws_order) 
        {
            Worksheet s = workbook.Worksheets[ws_order];
            return s;
        }

        // Get and return range to stores in Range variable.
        // wsheet --> Sheet instantiated of a Workbook variable, range1 and range2 --> strings in a Excel cell format. Ex.: "A1". 
        public Range GetRange(Worksheet wsheet, string range1, string range2) 
        {
            Range r = wsheet.Range[range1 + ":" + range2];
            return r;
        }

        // Returns the last line with data in worksheet instantied.
        // worksheet --> Sheet instantied, column --> A letter corresponding a colunm in sheet. Ex.: "A"
        public int GetLastLineData(Worksheet worksheet, string column)
        {
            int line_base = sheet_base.Range[column + "1048576"].End[XlDirection.xlUp].Row;
        }
    }
}