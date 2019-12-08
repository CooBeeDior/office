using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelTypeFormater
    {
        Action<ExcelRangeBase, object> SetHeaderCell();
        Action<ExcelRangeBase, object> SetBodyCell();

    }
}
