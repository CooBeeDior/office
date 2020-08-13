using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public class ExcelExcelPackageBuilder
    {
        public static ExcelPackage CreateExcelPackage()
        {
            return new ExcelPackage();
        }


        public static ExcelPackage CreateExcelPackage(Stream sm)
        {
            return new ExcelPackage(sm);
        }


        public static ExcelPackage CreateExcelPackage(byte[] buffer)
        {
            return new ExcelPackage(new MemoryStream(buffer));
        }

        public static ExcelPackage CreateExcelPackage(string filename)
        {
            return new ExcelPackage(new FileInfo(filename));
        }
    }
}
