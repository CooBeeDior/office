using CExcel.Service;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.DrawingCore;

namespace NpoiExcel.Service
{
    public class NpoiExcelExportFormater : IExcelExportFormater<ICell>
    {
        private ICellStyle headerCellStyle = null;
        private ICellStyle bodyCellStyle = null;


        public virtual Action<ICell, object> SetBodyCell()
        {
            return (c, o) =>
            {
                if (bodyCellStyle == null)
                {
                    bodyCellStyle = c.Sheet.Workbook.CreateCellStyle();

                    #region 设置单元格对齐方式   
                    bodyCellStyle.Alignment = HorizontalAlignment.Center;//水平居中
                    bodyCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                    #endregion

                    bodyCellStyle.BorderBottom = BorderStyle.Thin;//边框 
                }
                c.CellStyle = bodyCellStyle;         

                //设置值
                c.SetCellValue(o?.ToString());
            };
        }


        public virtual Action<ICell, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                if (headerCellStyle == null)
                {
                    headerCellStyle = c.Sheet.Workbook.CreateCellStyle();


                    #region 设置填充
                    headerCellStyle.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle.FillForegroundColor = IndexedColors.Green.Index;
                    #endregion

                    #region 设置单元格对齐方式   
                    headerCellStyle.Alignment = HorizontalAlignment.Center;//水平居中
                    headerCellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                    #endregion

                    headerCellStyle.BorderBottom = BorderStyle.Thin;//边框

                    #region 设置单元格字体样式
                    var font = c.Sheet.Workbook.CreateFont();
                    font.IsBold = true;//字体为粗体
                    font.Color = IndexedColors.White.Index;  //字体颜色
                    font.FontName = "微软雅黑";//字体
                    font.FontHeight = 12;//字体大小

                    headerCellStyle.SetFont(font);

                    #endregion            
                }
                c.CellStyle = headerCellStyle;
 
                c.SetCellValue(o?.ToString());
            };
        }
    }


}
