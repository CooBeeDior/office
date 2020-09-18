using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service.Impl
{
    public class ExcelProvider : IExcelProvider<ExcelPackage>
    {
        public IExcelExportService<ExcelPackage> CreateExcelExportService()
        {
            return new ExcelExportService();
        }

        public IExcelImportService<ExcelPackage> CreateExcelImportService()
        {
            return new ExcelImportService();
        }
    }
}
