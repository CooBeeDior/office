using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelImportService<TWorkbook>
    {
        IList<T> Import<T>(TWorkbook workbook, string sheetName) where T : class, new();
    }
}
