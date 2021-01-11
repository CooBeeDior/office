using CExcel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// excel创建者 只支持07之后的excel
    /// </summary>
    public class ExcelPackageBuilder : IWorkbookBuilder<ExcelPackage>
    {
        public ExcelPackage CreateWorkbook(CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            return new ExcelPackage();
        }


        public ExcelPackage CreateWorkbook(Stream sm, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            //
            return new ExcelPackage(sm);
        }


        public ExcelPackage CreateWorkbook(byte[] buffer, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            return new ExcelPackage(new MemoryStream(buffer));
        }

        public ExcelPackage CreateWorkbook(string filename, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            return new ExcelPackage(new FileInfo(filename));
        }
    }
}
