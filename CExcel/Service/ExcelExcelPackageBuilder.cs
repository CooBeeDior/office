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
    }
}
