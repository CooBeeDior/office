using CExcel.Service;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SpireExcel
{
    public class SpireWorkbookBuilder : IWorkbookBuilder<Workbook>
    {
        private readonly ExcelVersion excelVersion = ExcelVersion.Version2016;
        public Workbook CreateWorkbook()
        {
            var workbook = new Workbook() { Version = excelVersion };
            for (int i = workbook.Worksheets.Count-1; i >=1 ; i--)
            {
                workbook.Worksheets.RemoveAt(i);
            }
            return workbook;

        }

        public Workbook CreateWorkbook(Stream sm)
        {
            var workbook = CreateWorkbook();
            workbook.LoadFromStream(sm);
            return workbook;
        }

        public Workbook CreateWorkbook(byte[] buffer)
        {
            var workbook = CreateWorkbook();
            workbook.LoadFromStream(new MemoryStream(buffer));
            return workbook;
        }

        public Workbook CreateWorkbook(string filename)
        {
            var workbook = CreateWorkbook();
            workbook.LoadFromFile(filename);
            return workbook;
        }
    }
}
