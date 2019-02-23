using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

using _Excel = Microsoft.Office.Interop.Excel;
namespace ExcelTutorial
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook workBook;
        Worksheet workSheet;

        public Excel()
        {

        }

        public Excel(string path, int sheet)
        {
            this.path = path;
            workBook = excel.Workbooks.Open(path);
            workSheet = workBook.Worksheets[sheet];
        }

        public string ReadCell(int row, int col)
        {
            row++;
            col++;

            if(workSheet.Cells[row,col].Value2 != null)
            {
                return workSheet.Cells[row, col].Value2;
            }
            else
            {
                return "";
            }

        }

        public void WriteToCell(int row, int col, string text)
        {
            workSheet.Cells[++row, ++col].Value2 = text; 
        }

        public void Save()
        {
            workBook.Save();
        }

        public void SaveAs(string pathTo)
        {
            workBook.SaveAs(pathTo);
        }
        public void Close()
        {
            workBook.Close();
        }

        public void CreateNewFile()
        {
            workBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            workSheet = workBook.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet temptsheet = workBook.Worksheets.Add(After: workSheet);
        }

        public void SelectWorkSheet(int sheetNumber)
        {
            workSheet = workBook.Worksheets[sheetNumber];
        }

        public void DeleteWorkSheet(int sheetNumber)
        {
            workBook.Worksheets[sheetNumber].Delete();
        }

        public void ProtectSheet()
        {
            workSheet.Protect();
        }
        
        public void ProtectSheet(string password)
        {
            workSheet.Protect(password);
        }

        public void UnProtectSheet()
        {
            workSheet.Unprotect();
        }

        public void UnProtectSheet(string password)
        {
            workSheet.Unprotect(password);
        }

    }
}
