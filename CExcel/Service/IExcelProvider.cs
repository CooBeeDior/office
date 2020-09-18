using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelProvider<TWorkbook>  
    {
        IExcelExportService<TWorkbook> CreateExcelExportService();


        IExcelImportService<TWorkbook> CreateExcelImportService();
    }

   
}
