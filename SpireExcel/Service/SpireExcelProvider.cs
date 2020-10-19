using CExcel.Service;
using Spire.Xls;
using System;

namespace SpireExcel
{
    public class SpireExcelProvider : IExcelProvider<Workbook>
    {
        private readonly IExcelExportService<Workbook> _excelExportService;
        private readonly IExcelImportService<Workbook> _excelImportService;
        public SpireExcelProvider(IExcelExportService<Workbook> excelExportService, IExcelImportService<Workbook> excelImportService)
        {
            _excelExportService = excelExportService;
            _excelImportService = excelImportService;
        }

        public IExcelExportService<Workbook> CreateExcelExportService()
        {
            return _excelExportService;
        }


        public IExcelImportService<Workbook> CreateExcelImportService()
        {
            return _excelImportService;
        }
    }
}
