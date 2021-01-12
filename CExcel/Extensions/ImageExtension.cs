using OfficeOpenXml;
using System;
using System.Drawing;

namespace CExcel.Extensions
{
    public static class ImageExtension
    {
        #region private
        /// <summary>
        /// 获取自适应调整后的图片尺寸
        /// </summary>
        /// <param name="image"></param>
        /// <param name="cellColumnWidthInPix"></param>
        /// <param name="cellRowHeightInPix"></param>
        /// <returns>item1:调整后的图片宽度; item2:调整后的图片高度</returns>
        public static Tuple<int, int> GetAdjustImageSize(this Image image, int cellColumnWidthInPix, int cellRowHeightInPix)
        {
            int imageWidthInPix = image.Width;
            int imageHeightInPix = image.Height;
            //调整图片尺寸,适应单元格
            int adjustImageWidthInPix;
            int adjustImageHeightInPix;
            if (imageHeightInPix * cellColumnWidthInPix > imageWidthInPix * cellRowHeightInPix)
            {
                //图片高度固定,宽度自适应
                adjustImageHeightInPix = cellRowHeightInPix;
                double ratio = (1.0) * adjustImageHeightInPix / imageHeightInPix;
                adjustImageWidthInPix = (int)(imageWidthInPix * ratio);
            }
            else
            {
                //图片宽度固定,高度自适应
                adjustImageWidthInPix = cellColumnWidthInPix;
                double ratio = (1.0) * adjustImageWidthInPix / imageWidthInPix;
                adjustImageHeightInPix = (int)(imageHeightInPix * ratio);
            }
            return new Tuple<int, int>(adjustImageWidthInPix, adjustImageHeightInPix);
        }



        /// <summary>
        /// MeasureString
        /// </summary>
        /// <param name="s"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static float MeasureString(this string s, Font font)
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                return g.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
            }
        }






        #endregion
    }

}
