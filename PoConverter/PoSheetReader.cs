using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace PoConverter
{
    public static class PoSheetReader
    {
        public static PoSheet Read(string filePath)
        {
            var file = new PoSheet();

            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            var lines = xlRange.Rows;
            for (int i = PoSheet.StartLine; i < lines.Count; i++)
            {
                var currentRecord = new PoRecord() 
                { 
                    msgid = lines[i,1].Value,
                    msgstr = lines[i,2].Value
                };
                file.AddRecord(currentRecord);

            }
            return file;
        }
    }
}
