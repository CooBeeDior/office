using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// excel创建者
    /// </summary>
    public class ExcelPackageBuilder : IWorkbookBuilder<ExcelPackage>
    {
        public ExcelPackage CreateWorkbook()
        {
            return new ExcelPackage();
        }


        public ExcelPackage CreateWorkbook(Stream sm)
        {
            return new ExcelPackage(sm);
        }


        public ExcelPackage CreateWorkbook(byte[] buffer)
        {
            return new ExcelPackage(new MemoryStream(buffer));
        }

        public ExcelPackage CreateWorkbook(string filename)
        {
            return new ExcelPackage(new FileInfo(filename));
        }
    }
}
