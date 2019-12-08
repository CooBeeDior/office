using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace CExcel.Service.Impl
{
    public class DefaultExcelTypeFormater : IExcelTypeFormater
    {
        public virtual Action<ExcelRangeBase, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                c.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                c.Value = o;
            };
        }

        public virtual Action<ExcelRangeBase, object> SetBodyCell()
        {
            return (c, o) =>
            {
                c.Value = o;
            };
        }

      
    }
}
