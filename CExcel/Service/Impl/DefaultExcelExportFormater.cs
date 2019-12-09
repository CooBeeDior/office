using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace CExcel.Service.Impl
{
    public class DefaultExcelExportFormater : IExcelExportFormater
    {
        public virtual Action<ExcelRangeBase, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.Green);
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
