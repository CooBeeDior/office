using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
using CExcel.Models;
using CExcel.Service;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace SpireExcel.Extensions
{
    public static class ExcelExtension
    {
        public static Workbook AddSheet<T>(this Workbook workbook, IList<T> data = null) where T : class, new()
        {
            string sheetName = null;
            IExcelTypeFormater<Worksheet> defaultExcelTypeFormater = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (excelAttribute == null)
            {
                sheetName = typeof(T).Name;
                defaultExcelTypeFormater = new SpireExcelTypeFormater();
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
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<Worksheet>;
                }
                if (defaultExcelTypeFormater == null)
                {
                    defaultExcelTypeFormater = new SpireExcelTypeFormater();
                }
            }
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();


            IList<IExcelExportFormater<CellRange>> excelTypes = new List<IExcelExportFormater<CellRange>>();
            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
            int column = 1;

            //表头行
            foreach (var item in mainPropertieList)
            {
                IExcelExportFormater<CellRange> excelType = null;
                if (item.Value.ExportExcelType != null)
                {
                    excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ExportExcelType.FullName).FirstOrDefault();
                    if (excelType == null)
                    {
                        excelType = Activator.CreateInstance(item.Value.ExportExcelType) as IExcelExportFormater<CellRange>;
                        excelTypes.Add(excelType);
                    }
                }
                else
                {
                    excelType = defaultExcelExportFormater;
                }
                excelType.SetHeaderCell()?.Invoke(sheet[row, column], item.Value.Name);
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
                        IExcelExportFormater<CellRange> excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater<CellRange>;
                                excelTypes.Add(excelType);
                            }
                        }
                        else
                        {
                            excelType = defaultExcelExportFormater;
                        }
                        excelType.SetBodyCell()?.Invoke(sheet[row, column], mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return workbook;
        }
        public static Workbook AddSheet(this Workbook workbook, DataTable data)
        {
            string sheetName = data.TableName;
            IExcelTypeFormater<Worksheet> defaultExcelTypeFormater = new SpireExcelTypeFormater();

            Worksheet sheet = workbook.Worksheets[sheetName];
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

            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
            int column = 1;

            //表头行
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(sheet[row, column], headerName);
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
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return workbook;

        }
        public static Workbook AddSheetHeader(this Workbook workbook, string sheetName, IList<SpireHeaderInfo> headers)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }
            IExcelTypeFormater<Worksheet> defaultExcelTypeFormater = new SpireExcelTypeFormater();

            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);


            IList<IExcelExportFormater<CellRange>> excelTypes = new List<IExcelExportFormater<CellRange>>();
            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(sheet[row, column], item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(sheet[row, column], item.HeaderName);
                }
                column++;
            }

            row++;


            return workbook;

        }

        public static Workbook AddBody(this Workbook workbook, string sheetName, IList<IList<object>> data)
        {
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
                int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item is ExportCellValue<CellRange> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(sheet[row, column], cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], cellValue.Value);
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
                            var formatterPropertyInfo = item.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<CellRange>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(value) as IExcelExportFormater<CellRange>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(sheet[row, column], value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return workbook;

        }

        public static Workbook AddBody(this Workbook workbook, string sheetName, IList<IDictionary<string, object>> data)
        {
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            if (data != null && data.Any())
            {
                IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
                int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        if (item.Value is ExportCellValue<CellRange> cellValue)
                        {
                            if (cellValue?.ExportFormater != null)
                            {
                                cellValue?.ExportFormater.SetBodyCell()?.Invoke(sheet[row, column], cellValue.Value);
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], cellValue.Value);
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
                            var formatterPropertyInfo = item.Value.GetType().GetProperties().Where(o => typeof(IExcelExportFormater<CellRange>).IsAssignableFrom(o.PropertyType)).FirstOrDefault();
                            if (formatterPropertyInfo != null)
                            {
                                var formatterValue = formatterPropertyInfo.GetValue(item.Value) as IExcelExportFormater<CellRange>;
                                if (formatterValue != null)
                                {
                                    formatterValue.SetBodyCell()?.Invoke(sheet[row, column], value);
                                }
                                else
                                {
                                    defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], value);
                                }
                            }
                            else
                            {
                                defaultExcelExportFormater.SetBodyCell()?.Invoke(sheet[row, column], value);
                            }

                        }


                        column++;
                    }

                    row++;
                }
            }
            return workbook;

        }


        public static Workbook AddBody<T>(this Workbook workbook, IList<T> data, string sheetName = null) where T : class, new()
        {

            IExcelTypeFormater<Worksheet> defaultExcelTypeFormater = null;
            var excelAttribute = typeof(T).GetCustomAttribute<ExcelAttribute>();
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                if (excelAttribute == null)
                {
                    defaultExcelTypeFormater = new SpireExcelTypeFormater();
                }
                else
                {
                    if (excelAttribute.ExportExcelType != null)
                    {
                        defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<Worksheet>;
                    }
                    if (defaultExcelTypeFormater == null)
                    {
                        defaultExcelTypeFormater = new SpireExcelTypeFormater();
                    }
                }

            }
            else
            {
                if (excelAttribute == null)
                {
                    sheetName = typeof(T).Name;
                    defaultExcelTypeFormater = new SpireExcelTypeFormater();
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
                            sheetName = workbook.Worksheets.LastOrDefault().Name;
                        }

                    }
                    else
                    {
                        sheetName = excelAttribute.SheetName;
                    }
                    if (excelAttribute.ExportExcelType != null)
                    {
                        defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<Worksheet>;
                    }
                    if (defaultExcelTypeFormater == null)
                    {
                        defaultExcelTypeFormater = new SpireExcelTypeFormater();
                    }
                }
            }

            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(sheet);

            var mainPropertieList = typeof(T).ToColumnDic();


            IList<IExcelExportFormater<CellRange>> excelTypes = new List<IExcelExportFormater<CellRange>>();
            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (sheet.LastDataRow == -1 ? 0 : sheet.LastDataRow) + 1;
            int column = 1;



            //数据行 
            if (data != null && data.Any())
            {
                foreach (var item in data)
                {
                    column = 1;
                    foreach (var mainPropertie in mainPropertieList)
                    {
                        IExcelExportFormater<CellRange> excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater<CellRange>;
                                excelTypes.Add(excelType);
                            }
                        }
                        else
                        {
                            excelType = defaultExcelExportFormater;
                        }
                        excelType.SetBodyCell()?.Invoke(sheet[row, column], mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return workbook;
        }

        public static Workbook AddErrors(this Workbook workbook, string sheetName, IList<ExportExcelError> errors, Action<CellRange, string> action = null)
        {
            if (errors == null || !errors.Any())
            {
                return workbook;
            }
            var workSheet = workbook.Worksheets[sheetName];
            if (workSheet == null)
            {
                throw new Exception($"{sheetName}不存在");
            }
            if (action == null)
            {
                action = (cell, msg) =>
                {
                    cell.Style.Color = Color.Red;
                    if (cell.Comment == null)
                    {
                        cell.AddComment().Text = msg;
                    }
                    else
                    {
                        cell.Comment.Text = msg;
                    }
                };
            }

            foreach (var item in errors)
            {
                var cell = workSheet[item.Row, item.Column];
                action(cell, item.Message);
            }
            return workbook;

        }
        public static Workbook AddErrors<T>(this Workbook ep, IList<ExportExcelError> errors, Action<CellRange, string> action = null)
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
    }
}
