using CExcel.Service;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NpoiExcel.Service
{ 
    public class NpoiExcelProvider : IExcelProvider<IWorkbook>
    {
        private readonly IExcelExportService<IWorkbook> _excelExportService;
        private readonly IExcelImportService<IWorkbook> _excelImportService;
        public NpoiExcelProvider(IExcelExportService<IWorkbook> excelExportService, IExcelImportService<IWorkbook> excelImportService)
        {
            _excelExportService = excelExportService;
            _excelImportService = excelImportService;
        }

        public IExcelExportService<IWorkbook> CreateExcelExportService()
        {
            return _excelExportService;
        }


        public IExcelImportService<IWorkbook> CreateExcelImportService()
        {
            return _excelImportService;
        }
    }
}
