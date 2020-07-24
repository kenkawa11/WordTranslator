using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace WordTranslator
{
    class Excel : IDisposable
    {
        public Dictionary<string, Worksheet> Sheets = new Dictionary<string, Worksheet>();
        private Application excel;
        private Workbook workbook;
        public Excel(string filePath)
        {
            excel = new Application { Visible = false };
            workbook = excel.Workbooks.Open(filePath);

            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                Worksheet sheet = workbook.Sheets.Item[i + 1];
                Sheets.Add(sheet.Name, sheet);
            }
        }
        public void Dispose()
        {
            workbook.Close();
            workbook = null;
            excel.Quit();
            excel = null;
        }
        public int ReadCellInt(string sheetKey, string address)
        {
            var value = ReadCell(sheetKey, address);
            var integer = (int)(Math.Round(float.Parse(value)));
            return integer;
        }
        public string ReadCell(string sheetKey, string address)
        {
            var range = Sheets[sheetKey].get_Range(address);
            var cell = range.Cells[1, 1];
            return cell.Value.ToString();
        }
    }
}
