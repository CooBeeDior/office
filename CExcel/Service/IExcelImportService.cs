using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    /// <summary>
    /// 导入服务
    /// </summary>
    /// <typeparam name="TWorkbook"></typeparam>
    public interface IExcelImportService<TWorkbook>
    {
        IList<T> Import<T>(TWorkbook workbook, string sheetName = null) where T : class, new();
    }
}
