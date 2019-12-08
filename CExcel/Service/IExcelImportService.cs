using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelImportService<TWorkbook>
    {
        IList<T> Export<T>(TWorkbook workbook, string sheetName) where T : class, new();
    }
}
