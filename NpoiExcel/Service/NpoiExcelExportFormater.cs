﻿using CExcel.Service;
using NPOI.SS.UserModel;
using System;
using System.Drawing;
namespace NpoiExcel.Service
{
    public class NpoiExcelExportFormater : IExcelExportFormater<ICell>
    {
        public virtual Action<ICell, object> SetBodyCell()
        {
            return (c, o) =>
            {
                #region 设置单元格对齐方式   
                c.CellStyle.Alignment = HorizontalAlignment.Center;//水平居中
                c.CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                #endregion


                c.SetAsActiveCell();//单元格的宽度
                c.CellStyle.BorderBottom = BorderStyle.Thin;//边框


                //设置值
                c.SetCellValue(o?.ToString());
            };
        }

        public virtual Action<ICell, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                #region 设置单元格对齐方式   
                c.CellStyle.Alignment = HorizontalAlignment.Center;//水平居中
                c.CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                #endregion

                c.SetAsActiveCell();//单元格的宽度

                c.CellStyle.BorderBottom = BorderStyle.Thin;//边框

                #region 设置单元格字体样式
                var font = c.Sheet.Workbook.CreateFont();
                font.IsBold = true;//字体为粗体
                font.Color = (short)Color.White.ToArgb();//字体颜色
                font.FontName = "微软雅黑";//字体
                font.FontHeight = 12;//字体大小

                c.CellStyle.SetFont(font);

                #endregion

                c.CellStyle.FillBackgroundColor = (short)Color.Green.ToArgb();

                c.SetCellValue(o?.ToString());
            };
        }
    }


}
