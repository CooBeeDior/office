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
                if (excelAttribute.ExportExcelType != null)
                {
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<ISheet>;
                }
                else
                {
                    defaultExcelTypeFormater = new NpoiExcelTypeFormater();
                }
            }
            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }

            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();

            IList<IExcelExportFormater<ICell>> excelTypes = new List<IExcelExportFormater<ICell>>();
            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.LastRowNum ?? 0) + 1;
            int column = 1;

            //表头行
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
                excelType.SetHeaderCell()?.Invoke(sheet.GetRow(row).GetCell(column), item.Value.Name);
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
                        excelType.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), mainValue);
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

            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var headerNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerNames.Add(data.Columns[i].ColumnName);
            }
            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.LastRowNum ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(sheet.GetRow(row).GetCell(column), headerName);
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
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), mainValue);
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

            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.LastRowNum ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(sheet.GetRow(row).GetCell(column), item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(sheet.GetRow(row).GetCell(column), item.HeaderName);
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

            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
            int row = (sheet?.LastRowNum ?? 0) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(sheet.GetRow(row).GetCell(column), item);
                }
                else
                {
                    action.Invoke(sheet.GetRow(row).GetCell(column), item);
                }

                column++;
            }

            row++;


            return workbook;

        }


        public static IWorkbook AddBody(this IWorkbook workbook, string sheetName, IList<IList<object>> data)
        {
            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
                int row = (sheet?.LastRowNum ?? 0) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item is ExportCellValue<ICell> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), cellValue.Value);
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
                                    formatterValue.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
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
            ISheet sheet= workbook.GetSheet(sheetName);
            if (sheet == null)
            {
               sheet= workbook.CreateSheet(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<ICell> defaultExcelExportFormater = new NpoiExcelExportFormater();
                int row = (sheet?.LastRowNum ?? 0) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item.Value is ExportCellValue<ICell> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), cellValue.Value);
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
                                    formatterValue.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet.GetRow(row).GetCell(column), value);
                            }

                        }


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
                    cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    cell.CellStyle.FillBackgroundColor = (short)Color.Red.ToArgb();
                    if (workbook is HSSFWorkbook)
                    {
                        cell.CellComment.String = new HSSFRichTextString(msg);
                    }
                    else
                    {
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
            return cell.StringCellValue;
        }
    }
}
