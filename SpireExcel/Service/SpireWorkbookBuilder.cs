using CExcel.Models;
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

        public Workbook CreateWorkbook(CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            ExcelVersion version = ExcelVersion.Version2016;
            if (excelVersion == CExcelVersion.Version2003)
            {
                version = ExcelVersion.Version97to2003;
            }
            var workbook = new Workbook() { Version = version };
            workbook.Worksheets.Clear();
            return workbook;

        }

        public Workbook CreateWorkbook(Stream sm, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            var workbook = CreateWorkbook(excelVersion);
            workbook.LoadFromStream(sm);
            return workbook;
        }

        public Workbook CreateWorkbook(byte[] buffer, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            var workbook = CreateWorkbook(excelVersion);
            workbook.LoadFromStream(new MemoryStream(buffer));
            return workbook;
        }

        public Workbook CreateWorkbook(string filename, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            var workbook = CreateWorkbook(excelVersion);
            workbook.LoadFromFile(filename);
            return workbook;
        }
    }
}
