using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    /// <summary>
    /// excel提供者
    /// </summary>
    /// <typeparam name="TWorkbook"></typeparam>
    public interface IExcelProvider<TWorkbook>
    {
        IExcelExportService<TWorkbook> CreateExcelExportService();


        IExcelImportService<TWorkbook> CreateExcelImportService();
    }




}
