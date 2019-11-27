using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataReaderLib
{
    public class ExcelDataReader
    {
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        private Excel.Range xlRange;
        public ExcelDataReader()
        {
            Excel.Application xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Anil\Desktop\ExcelReaderLib\Test.xlsx");
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
        }

        public string GetExcelData(string Name)
        {
            Excel.Range colRange = xlWorksheet.Columns["$A:$A"];
            string[] secondColRange = xlRange.Find("Age").Cells.Address.Split(new char[] { '$' });
            Excel.Range _FindRange = colRange.Find(Name);
            if(_FindRange == null)
            {
                return "Name not found";
            }
            else
            {
                string[] rowAdd = _FindRange.Cells.AddressLocal.Split(new char[] { '$' });
                //var result = rowAdd[2];
                var result = xlWorksheet.Cells[rowAdd[2],secondColRange[1]];
                return result.ToString() ;
            }


        }
    }
}
