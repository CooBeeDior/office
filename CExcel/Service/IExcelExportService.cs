using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    /// <summary>
    /// 导出服务
    /// </summary>
    /// <typeparam name="TWorkbook"></typeparam>
    public interface IExcelExportService<TWorkbook>
    {
        TWorkbook Export<T>(IList<T> data = null) where T : class, new();

        TWorkbook Export(IList<object> data);
        Stream ToStream(TWorkbook workbook);
    }
}
