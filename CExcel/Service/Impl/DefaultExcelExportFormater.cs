using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// 导出格式化
    /// </summary>
    public class DefaultExcelExportFormater : IExcelExportFormater
    {

        public virtual Action<ExcelRangeBase, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                #region 设置单元格对齐方式   
                c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                c.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                #endregion

                #region 设置单元格字体样式
                c.Style.Font.Bold = true;//字体为粗体
                c.Style.Font.Color.SetColor(Color.White);//字体颜色
                c.Style.Font.Name = "微软雅黑";//字体
                c.Style.Font.Size = 12;//字体大小
                #endregion

                #region 设置单元格背景样式
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.Green);
                #endregion

                #region 设置单元格边框
                c.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));//设置单元格所有边框
                c.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;//单独设置单元格底部边框样式和颜色（上下左右均可分开设置）
                c.Style.Border.Bottom.Color.SetColor(Color.FromArgb(191, 191, 191));
                #endregion

                //c.Style.Numberformat.Format = "#,##0.00";//这是保留两位小数

                //设置值
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
