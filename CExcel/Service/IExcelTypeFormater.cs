using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    /// <summary>
    /// excel格式化
    /// </summary>
    public interface IExcelTypeFormater<TWorksheet>
    {
        Action<TWorksheet> SetExcelWorksheet();

    }
}
