using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    /// <summary>
    /// 导出格式化
    /// </summary>
    public interface IExcelExportFormater
    { 
        Action<ExcelRangeBase, object> SetHeaderCell();
        Action<ExcelRangeBase, object> SetBodyCell();

    }
}
