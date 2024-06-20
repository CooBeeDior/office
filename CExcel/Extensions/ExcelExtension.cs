using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Models;
using CExcel.Service;
using CExcel.Service.Impl;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace CExcel.Extensions
{

    public static class ExcelExtension
    {
        public static ExcelPackage AddSheet<T>(this ExcelPackage ep, IList<T> data = null) where T : class, new()
        {
            ExcelWorkbook workbook = ep.Workbook;
            string sheetName = null;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (excelAttribute == null)
            {
                sheetName = typeof(T).Name;
                defaultExcelTypeFormater = new DefaultExcelTypeFormater();
            }
            else
            {
                if (excelAttribute.IsIncrease)
                {
                    if (workbook.Worksheets.Count == 0)
                    {
                        sheetName = $"{excelAttribute.SheetName}";
                    }
                    else
                    {
                        sheetName = $"{excelAttribute.SheetName}{workbook.Worksheets.Count}";
                    }

                }
                else
                {
                    sheetName = excelAttribute.SheetName;
                }
                if (excelAttribute.ExportExcelType != null)
                {
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<ExcelWorksheet>;
                }
                if (defaultExcelTypeFormater == null)
                {
                    defaultExcelTypeFormater = new DefaultExcelTypeFormater();
                }
            }
            ExcelWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }

            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();

            IList<IExcelExportFormater<ExcelRangeBase>> excelTypes = new List<IExcelExportFormater<ExcelRangeBase>>();
            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = (sheet?.Dimension?.Rows ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var item in mainPropertieList)
            {
                IExcelExportFormater<ExcelRangeBase> excelType = null;
                if (item.Value.ExportExcelType != null)
                {
                    excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ExportExcelType.FullName).FirstOrDefault();
                    if (excelType == null)
                    {
                        excelType = Activator.CreateInstance(item.Value.ExportExcelType) as IExcelExportFormater<ExcelRangeBase>;
                        excelTypes.Add(excelType);
                    }
                }
                else
                {
                    excelType = defaultExcelExportFormater;
                }
                excelType.SetHeaderCell()?.Invoke(sheet.Cells[row, column], item.Value.Name);
                column++;
            }

            row++;

            //数据行 
            if (data != null && data.Any())
            {
                foreach (var item in data)
                {
                    column = 1;
                    foreach (var mainPropertie in mainPropertieList)
                    {
                        IExcelExportFormater<ExcelRangeBase> excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater<ExcelRangeBase>;
                                excelTypes.Add(excelType);
                            }
                        }
                        else
                        {
                            excelType = defaultExcelExportFormater;
                        }
                        excelType.SetBodyCell()?.Invoke(sheet.Cells[row, column], mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return ep;

        }

        public static ExcelPackage AddSheet(this ExcelPackage ep, DataTable data)
        {
            ExcelWorkbook workbook = ep.Workbook;
            string sheetName = data.TableName;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var headerNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerNames.Add(data.Columns[i].ColumnName);
            }
            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = (sheet?.Dimension?.Rows ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(sheet.Cells[row, column], headerName);
                column++;
            }

            row++;

            //数据行 
            if (data != null && data.Rows.Count > 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    column = 1;
                    foreach (var headerName in headerNames)
                    {
                        var mainValue = data.Rows[i][headerName];
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.Cells[row, column], mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return ep;

        }
        public static ExcelPackage AddSheetHeader(this ExcelPackage ep, string sheetName, IList<HeaderInfo> headers)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }
            ExcelWorkbook workbook = ep.Workbook;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = (sheet?.Dimension?.Rows ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(sheet.Cells[row, column], item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(sheet.Cells[row, column], item.HeaderName);
                }
                column++;
            }

            row++;


            return ep;

        }
        public static ExcelPackage AddSheetHeader(this ExcelPackage ep, string sheetName, IList<string> headers, Action<ExcelRangeBase, object> action = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }
            ExcelWorkbook workbook = ep.Workbook;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = (sheet?.Dimension?.Rows ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(sheet.Cells[row, column], item);
                }
                else
                {
                    action.Invoke(sheet.Cells[row, column], item);
                }

                column++;
            }

            row++;


            return ep;

        }
        public static ExcelPackage AddBody(this ExcelPackage ep, string sheetName, IList<IList<object>> data)
        {
            ExcelWorkbook workbook = ep.Workbook;
            ExcelWorksheet ws = workbook.Worksheets[sheetName];
            if (ws == null)
            {
                ws = workbook.Worksheets.Add(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
                int row = (ws?.Dimension?.Rows ?? 0) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item is ExportCellValue<ExcelRangeBase> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], cellValue.Value);
                            }

                        }
                        else
                        {
                            var valuePropertyInfo = item.GetType().GetProperties().Where(o => o.Name.Equals("Value", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                            var value = valuePropertyInfo?.GetValue(item);
                            if (valuePropertyInfo == null || value == null)
                            {
                                throw new Exception("Value值不能为空");
                            }
                            var formatterPropertyInfo = item.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<ExcelRangeBase>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(item) as IExcelExportFormater<ExcelRangeBase>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return ep;

        }
        public static ExcelPackage AddBody(this ExcelPackage ep, string sheetName, IList<IDictionary<string, object>> data)
        {
            ExcelWorkbook workbook = ep.Workbook;
            ExcelWorksheet ws = workbook.Worksheets[sheetName];
            if (ws == null)
            {
                ws = workbook.Worksheets.Add(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
                int row = (ws?.Dimension?.Rows ?? 0) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item.Value is ExportCellValue<ExcelRangeBase> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], cellValue.Value);
                            }

                        }
                        else
                        {
                            var valuePropertyInfo = item.Value.GetType().GetProperties().Where(o => o.Name.Equals("Value", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                            var value = valuePropertyInfo?.GetValue(item.Value);
                            if (valuePropertyInfo == null || value == null)
                            {
                                throw new Exception("Value值不能为空");
                            }
                            var formatterPropertyInfo = item.Value.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<ExcelRangeBase>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(item.Value) as IExcelExportFormater<ExcelRangeBase>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return ep;

        }
        public static ExcelPackage AddErrors<T>(this ExcelPackage ep, IList<ExportExcelError> errors, Action<ExcelRangeBase, string> action = null)
        {
            string sheetName = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (excelAttribute != null)
            {
                sheetName = excelAttribute.SheetName;
            }
            else
            {
                sheetName = nameof(T);
            }
            return ep.AddErrors(sheetName, errors, action);
        }

        public static ExcelPackage AddErrors(this ExcelPackage ep, string sheetName, IList<ExportExcelError> errors, Action<ExcelRangeBase, string> action = null)
        {
            if (errors == null || !errors.Any())
            {
                return ep;
            }
            var workSheet = ep.Workbook.Worksheets[sheetName];
            if (workSheet == null)
            {
                throw new Exception($"{sheetName}不存在");
            }
            if (action == null)
            {
                action = (cell, msg) =>
                {
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    if (cell.Comment == null)
                    {
                        cell.AddComment(msg, "管理员");
                    }
                    else
                    {
                        cell.Comment.Text = msg;
                    }
                };
            }

            foreach (var item in errors)
            {
                var cell = workSheet.Cells[item.Row, item.Column];
                action(cell, item.Message);
            }
            return ep;

        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="imageBytes"></param>
        /// <param name="rowNum"></param>
        /// <param name="columnNum"></param>
        /// <param name="autofit"></param>
        public static void AddPicture(this ExcelWorksheet worksheet, byte[] imageBytes, int rowNum, int columnNum, bool autofit = true)
        {
            AddPicture(worksheet, imageBytes, worksheet.Cells[rowNum, columnNum], autofit);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="imageBytes"></param>
        /// <param name="rowNum"></param>
        /// <param name="columnNum"></param>
        /// <param name="autofit"></param>
        public static void AddPicture(this ExcelWorksheet worksheet, byte[] imageBytes, ExcelRangeBase cell, bool autofit)
        {
            using (var image = Image.FromStream(new MemoryStream(imageBytes)))
            {
                var picture = worksheet.Drawings.AddPicture($"image_{DateTime.Now.Ticks}", image);
                int cellColumnWidthInPix = GetWidthInPixels(cell);
                int cellRowHeightInPix = GetHeightInPixels(cell);
                int adjustImageWidthInPix = cellColumnWidthInPix;
                int adjustImageHeightInPix = cellRowHeightInPix;
                if (autofit)
                {
                    //图片尺寸适应单元格
                    var adjustImageSize = image.GetAdjustImageSize(cellColumnWidthInPix, cellRowHeightInPix);
                    adjustImageWidthInPix = adjustImageSize.Item1;
                    adjustImageHeightInPix = adjustImageSize.Item2;
                }
                //设置为居中显示
                int columnOffsetPixels = (int)((cellColumnWidthInPix - adjustImageWidthInPix) / 2.0);
                int rowOffsetPixels = (int)((cellRowHeightInPix - adjustImageHeightInPix) / 2.0);
                picture.SetSize(adjustImageWidthInPix, adjustImageHeightInPix);
                picture.SetPosition(cell.Start.Row - 1, rowOffsetPixels, cell.Start.Column - 1, columnOffsetPixels);
            }
        }

        public static List<KeyValuePair<PropertyInfo, ExcelColumnAttribute>> ToColumnDic(this Type @type)
        {
            Dictionary<PropertyInfo, ExcelColumnAttribute> mainDic = new Dictionary<PropertyInfo, ExcelColumnAttribute>();
            int order = 1;
            @type.GetProperties().ToList().ForEach(o =>
            {
                var attribute = o.GetCustomAttribute<ExcelColumnAttribute>();
                if (attribute == null)
                {
                    if (mainDic.Count > 0)
                    {
                        order = mainDic.ElementAt(mainDic.Count - 1).Value.Order + 1;
                    }
                    attribute = new ExcelColumnAttribute(o.Name, order);
                    mainDic.Add(o, attribute);
                }
                else if (!attribute.Ignore)
                {
                    if (mainDic.Count > 0 && attribute.Order == 0)
                    { 
                        order = mainDic.ElementAt(mainDic.Count - 1).Value.Order + 1;
                        attribute.Order = order;
                    }
                    mainDic.Add(o, attribute);
                }
            });

            var mainPropertieList = mainDic.OrderBy(o => o.Value.Order).ToList();
            return mainPropertieList;
        }

        public static string GetCellAddress(this Type @type, string name)
        {
            var mainPropertieList = type.ToColumnDic();

            int currentIndex = 0;
            foreach (var item in mainPropertieList)
            {
                if (item.Key.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                {
                    return currentIndex.IndexToAddress();
                }
                currentIndex++;
            }
            return null;
        }

        public static string IndexToAddress(this int index)
        {
            string[] columnAddress = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            if (index < columnAddress.Length)
            {
                return columnAddress[index];
            }
            if (index >= columnAddress.Length)
            {
                int currentIndex = columnAddress.Length;
                for (int i = 0; i < columnAddress.Length; i++)
                {
                    for (int j = 0; i < columnAddress.Length; j++)
                    {
                        var address = $"{columnAddress[i]}{columnAddress[j]}";
                        if (index == currentIndex)
                        {
                            return address;
                        }
                        currentIndex++;
                    }


                }

            }

            throw new Exception("定义字段过多");
        }

        #region
        /// <summary>
        /// 获取单元格的宽度(像素)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static int GetWidthInPixels(ExcelRangeBase cell)
        {
            double columnWidth = cell.Worksheet.Column(cell.Start.Column).Width;
            Font font = new Font(cell.Style.Font.Name, cell.Style.Font.Size, FontStyle.Regular);
            double pxBaseline = Math.Round("1234567890".MeasureString(font) / 10);
            return (int)(columnWidth * pxBaseline);
        }

        /// <summary>
        /// 获取单元格的高度(像素)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static int GetHeightInPixels(ExcelRangeBase cell)
        {
            double rowHeight = cell.Worksheet.Row(cell.Start.Row).Height;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiY = graphics.DpiY;
                return (int)(rowHeight * (1.0 / 70) * dpiY);
            }
        }
        #endregion
    }
}
