using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// excel提供者
    /// </summary>
    public class ExcelProvider : IExcelProvider<ExcelPackage>
    {
        private readonly IExcelExportService<ExcelPackage> _excelExportService;
        private readonly IExcelImportService<ExcelPackage> _excelImportService;
        public ExcelProvider(IExcelExportService<ExcelPackage> excelExportService, IExcelImportService<ExcelPackage> excelImportService)
        {
            _excelExportService = excelExportService;
            _excelImportService = excelImportService;
        }
        public IExcelExportService<ExcelPackage> CreateExcelExportService()
        {
            return _excelExportService;
        }

        public IExcelImportService<ExcelPackage> CreateExcelImportService()
        {
            return _excelImportService;
        }
    }
}
