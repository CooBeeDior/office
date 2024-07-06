using CExcel;
using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
using CExcel.Models;
using CExcel.Service;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NpoiExcel.Models;
using NpoiExcel.Service;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NpoiExcel.Extensions
{
    public static class ExcelExtension
    {
        public static IWorkbook AddSheet<T>(this IWorkbook workbook, IList<T> data = null) where T : class, new()
        {
            string sheetName = null;
            IExcelTypeFormater<ISheet> defaultExcelTypeFormater = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (excelAttribute == null)
            {
                sheetName = typeof(T).Name;
                defaultExcelTypeFormater = new NpoiExcelTypeFormater();
            }
            else
            {
                if (excelAttribute.IsIncrease)
                {
                    if (workbook.NumberOfSheets == 0)
                    {
                        sheetName = $"{excelAttribute.SheetName}";
                    }
                    else
                    {
                        sheetName = $"{excelAttribute.SheetName}{workbook.NumberOfSheets}";
                    }

                }
                else
                {
                    sheetName = excelAttribute.SheetName;
                }
                if (excelAttribute.ExportExcelType != null && typeof(IExcelTypeFormater<ISheet>).IsAssignableFrom(excelAttribute.ExportExcelType))
                {
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<ISheet>;
                }
                if (defaultExcelTypeFormater == null)
                {
                    defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                }
            }
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();

            IList<IExcelExportFormater<ICell>> excelTypes = new List<IExcelExportFormater<ICell>>();
            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.PhysicalNumberOfRows ?? 0);
            int column = 0;

            //表头行
            var headerRowCell = sheet.CreateRow(row);
            foreach (var item in mainPropertieList)
            {
                IExcelExportFormater<ICell> excelType = null;
                if (item.Value.ExportExcelType != null)
                {
                    excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ExportExcelType.FullName).FirstOrDefault();
                    if (excelType == null)
                    {
                        excelType = Activator.CreateInstance(item.Value.ExportExcelType) as IExcelExportFormater<ICell>;
                        excelTypes.Add(excelType);
                    }
                }
                else
                {
                    excelType = defaultExcelExportFormater;
                }
                excelType.SetHeaderCell()?.Invoke(headerRowCell.CreateCell(column), item.Value.Name);
                column++;
            }

            row++;

            //数据行 
            if (data != null && data.Any())
            {
                foreach (var item in data)
                {
                    var rowCell = sheet.CreateRow(row);
                    column = 0;
                    foreach (var mainPropertie in mainPropertieList)
                    {
                        IExcelExportFormater<ICell> excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater<ICell>;
                                excelTypes.Add(excelType);
                            }
                        }
                        else
                        {
                            excelType = defaultExcelExportFormater;
                        }
                        excelType.SetBodyCell()?.Invoke(rowCell.CreateCell(column), mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return workbook;

        }

        public static IWorkbook AddSheet(this IWorkbook workbook, DataTable data)
        {
            string sheetName = data.TableName;
            IExcelTypeFormater<ISheet> defaultExcelTypeFormater = new NpoiExcelTypeFormater();

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var headerNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerNames.Add(data.Columns[i].ColumnName);
            }
            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.PhysicalNumberOfRows ?? 0);
            int column = 0;

            //表头行
            var headerRowCell = sheet.CreateRow(row);
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(headerRowCell.CreateCell(column), headerName);
                column++;
            }

            row++;

            //数据行 
            if (data != null && data.Rows.Count > 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    var rowCell = sheet.CreateRow(row);
                    column = 0;
                    foreach (var headerName in headerNames)
                    {
                        var mainValue = data.Rows[i][headerName];
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return workbook;

        }


        public static IWorkbook AddSheetHeader(this IWorkbook workbook, string sheetName, IList<NpoiHeaderInfo> headers)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }
            IExcelTypeFormater<ISheet> defaultExcelTypeFormater = new NpoiExcelTypeFormater();

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.PhysicalNumberOfRows ?? 0);
            int column = 0;

            //表头行
            var headerRowCell = sheet.CreateRow(row);
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(headerRowCell.CreateCell(column), item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(headerRowCell.CreateCell(column), item.HeaderName);
                }
                column++;
            }

            row++;


            return workbook;

        }
        public static IWorkbook AddSheetHeader(this IWorkbook workbook, string sheetName, IList<string> headers, Action<ICell, object> action = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }
            IExcelTypeFormater<ISheet> defaultExcelTypeFormater = new NpoiExcelTypeFormater();

            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.PhysicalNumberOfRows ?? 0);
            int column = 0;

            //表头行
            var headerRowCell = sheet.CreateRow(row);
            foreach (var item in headers)
            {
                if (action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(headerRowCell.CreateCell(column), item);
                }
                else
                {
                    action.Invoke(headerRowCell.CreateCell(column), item);
                }

                column++;
            }

            row++;


            return workbook;

        }


        public static IWorkbook AddBody(this IWorkbook workbook, string sheetName, IList<IList<object>> data)
        {
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
                int row = (sheet?.PhysicalNumberOfRows ?? 0);
                foreach (var dic in data)
                {

                    int column = 0;
                    var rowCell = sheet.CreateRow(row);
                    foreach (var item in dic)
                    {
                        if (item is ExportCellValue<ICell> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), cellValue.Value);
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
                            var formatterPropertyInfo = item.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<ICell>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(item) as IExcelExportFormater<ICell>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return workbook;

        }

        public static IWorkbook AddBody(this IWorkbook workbook, string sheetName, IList<IDictionary<string, object>> data)
        {
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
                int row = (sheet?.PhysicalNumberOfRows ?? 0);
                foreach (var dic in data)
                {

                    int column = 0;
                    var rowCell = sheet.CreateRow(row);
                    foreach (var item in dic)
                    {
                        if (item.Value is ExportCellValue<ICell> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), cellValue.Value);
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
                            var formatterPropertyInfo = item.Value.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<ICell>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(item.Value) as IExcelExportFormater<ICell>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(rowCell.CreateCell(column), value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return workbook;

        }

        public static IWorkbook AddBody<T>(this IWorkbook workbook, IList<T> data, string sheetName = null) where T : class, new()
        {

            IExcelTypeFormater<ISheet> defaultExcelTypeFormater = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                if (excelAttribute == null)
                {
                    defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                }
                else
                {
                    if (excelAttribute.ExportExcelType != null && typeof(IExcelTypeFormater<ISheet>).IsAssignableFrom(excelAttribute.ExportExcelType))
                    {
                        defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<ISheet>;
                    }
                    if (defaultExcelTypeFormater == null)
                    {
                        defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                    }
                }
            }
            else
            {
                if (excelAttribute == null)
                {
                    sheetName = typeof(T).Name;
                    defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                }
                else
                {
                    if (excelAttribute.IsIncrease)
                    {
                        if (workbook.NumberOfSheets == 0)
                        {
                            sheetName = $"{excelAttribute.SheetName}";
                        }
                        else
                        {
                            sheetName = workbook.GetSheetAt(workbook.NumberOfSheets - 1).SheetName;
                        }

                    }
                    else
                    {
                        sheetName = excelAttribute.SheetName;
                    }


                    if (excelAttribute.ExportExcelType != null && typeof(IExcelTypeFormater<ISheet>).IsAssignableFrom(excelAttribute.ExportExcelType))
                    {
                        defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<ISheet>;
                    }
                    if (defaultExcelTypeFormater == null)
                    {
                        defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                    }
                }
            }



            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();

            IList<IExcelExportFormater<ICell>> excelTypes = new List<IExcelExportFormater<ICell>>();
            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.PhysicalNumberOfRows ?? 0);
            int column = 0;

            //数据行 
            if (data != null && data.Any())
            {
                foreach (var item in data)
                {
                    var rowCell = sheet.CreateRow(row);
                    column = 0;
                    foreach (var mainPropertie in mainPropertieList)
                    {
                        IExcelExportFormater<ICell> excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater<ICell>;
                                excelTypes.Add(excelType);
                            }
                        }
                        else
                        {
                            excelType = defaultExcelExportFormater;
                        }
                        excelType.SetBodyCell()?.Invoke(rowCell.CreateCell(column), mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return workbook;

        }
        public static IWorkbook AddErrors<T>(this IWorkbook workbook, IList<ExportExcelError> errors, Action<ICell, string> action = null)
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
            return workbook.AddErrors(sheetName, errors, action);
        }

        public static IWorkbook AddErrors(this IWorkbook workbook, string sheetName, IList<ExportExcelError> errors, Action<ICell, string> action = null)
        {
            if (errors == null || !errors.Any())
            {
                return workbook;
            }
            var workSheet = workbook.GetSheet(sheetName);
            if (workSheet == null)
            {
                throw new Exception($"{sheetName}不存在");
            }
            if (action == null)
            {


                action = (cell, msg) =>
                {
                    var cellStyle = cell.Sheet.Workbook.CreateCellStyle();
                    cellStyle.FillPattern = FillPattern.FineDots;
                    cellStyle.FillBackgroundColor = IndexedColors.Red.Index;


                    #region 设置单元格字体样式
                    var font = cell.Sheet.Workbook.CreateFont();
                    font.IsBold = true;//字体为粗体
                    font.Color = IndexedColors.White.Index;  //字体颜色
                    font.FontName = "微软雅黑";//字体
                    font.FontHeight = 12;//字体大小
                    cellStyle.SetFont(font);

                    cell.CellStyle = cellStyle;
                    #endregion   

                    if (cell is HSSFCell)
                    {
                        if (cell.CellComment == null)
                        {
                            // 创建绘图主控制器(用于包括单元格注释在内的所有形状的顶级容器)
                            IDrawing patriarch = cell.Sheet.CreateDrawingPatriarch();
                            // 客户端锚定定义工作表中注释的大小和位置
                            //(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2) 
                            //前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
                            IComment comment = patriarch.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 2, cell.RowIndex + 4));
                            comment.Author = "系统管理员";
                            cell.CellComment = comment;
                        }
                        cell.CellComment.String = new HSSFRichTextString(msg);
                    }
                    else
                    {
                        if (cell.CellComment == null)
                        {
                            // 创建绘图主控制器(用于包括单元格注释在内的所有形状的顶级容器)
                            IDrawing patriarch = cell.Sheet.CreateDrawingPatriarch();
                            // 客户端锚定定义工作表中注释的大小和位置
                            //(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2) 
                            //前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
                            IComment comment = patriarch.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 2, cell.RowIndex + 4));
                            comment.Author = "系统管理员";
                            cell.CellComment = comment;
                        }
                        cell.CellComment.String = new XSSFRichTextString(msg);
                    }

                };
            }

            foreach (var item in errors)
            {
                var cell = workSheet.GetRow(item.Row).GetCell(item.Column);
                action(cell, item.Message);
            }
            return workbook;

        }



        public static object ToValue(this ICell cell)
        {
            object tempValue = "";
            if (cell == null)
            {
                return tempValue;
            }
            switch (cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.Blank:
                    break;
                case NPOI.SS.UserModel.CellType.Boolean:
                    tempValue = cell.BooleanCellValue;
                    break;
                case NPOI.SS.UserModel.CellType.Error:
                    break;
                case NPOI.SS.UserModel.CellType.Formula:
                    NPOI.SS.UserModel.IFormulaEvaluator fe = NPOI.SS.UserModel.WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
                    var cellValue = fe.Evaluate(cell);
                    switch (cellValue.CellType)
                    {
                        case NPOI.SS.UserModel.CellType.Blank:
                            break;
                        case NPOI.SS.UserModel.CellType.Boolean:
                            tempValue = cellValue.BooleanValue;
                            break;
                        case NPOI.SS.UserModel.CellType.Error:
                            break;
                        case NPOI.SS.UserModel.CellType.Formula:
                            break;
                        case NPOI.SS.UserModel.CellType.Numeric:
                            tempValue = cellValue.NumberValue;
                            break;
                        case NPOI.SS.UserModel.CellType.String:
                            tempValue = cellValue.StringValue;
                            break;
                        case NPOI.SS.UserModel.CellType.Unknown:
                            break;
                        default:
                            break;
                    }
                    break;
                case NPOI.SS.UserModel.CellType.Numeric:

                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                    {
                        tempValue = cell.DateCellValue.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        tempValue = cell.NumericCellValue;
                    }
                    break;
                case NPOI.SS.UserModel.CellType.String:
                    tempValue = cell.StringCellValue.Trim();
                    break;
                case NPOI.SS.UserModel.CellType.Unknown:
                    break;
                default:
                    break;
            }
            return tempValue;
        }

        public static ICell AddPicture(this ICell cell, byte[] bytes)
        {
            int pictureIdx = cell.Sheet.Workbook.AddPicture(bytes, PictureType.JPEG);
            IDrawing patriarch = cell.Sheet.CreateDrawingPatriarch();
            // 插图片的位置  HSSFClientAnchor（dx1,dy1,dx2,dy2,col1,row1,col2,row2) 后面再作解释
            if (cell is HSSFCell)
            {
                HSSFClientAnchor anchor = new HSSFClientAnchor(70, 10, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 1, cell.RowIndex + 1);
                //把图片插到相应的位置
                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            }
            else
            {
                XSSFClientAnchor anchor = new XSSFClientAnchor(70, 10, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 1, cell.RowIndex + 1);
                //把图片插到相应的位置
                XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            }

            return cell;
        }

        public static ICell AddPicture(this ICell cell, MemoryStream stream)
        {
            var buffer = new byte[stream.Length];
            stream.Write(buffer, 0, buffer.Length);
            stream.Seek(0, SeekOrigin.Begin);
            return cell.AddPicture(buffer);
        }
    }
}
