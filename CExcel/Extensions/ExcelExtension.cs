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
using System.Linq;
using System.Reflection;
using System.Text;

namespace CExcel.Extensions
{

    public static class ExcelExtension
    {
        public static ExcelPackage AddSheet<T>(this ExcelPackage ep, IList<T> data = null) where T : class, new()
        {
            ExcelWorkbook wb = ep.Workbook;
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
                    if (wb.Worksheets.Count == 0)
                    {
                        sheetName = $"{excelAttribute.SheetName}";
                    }
                    else
                    {
                        sheetName = $"{excelAttribute.SheetName}{wb.Worksheets.Count}";
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
                else
                {
                    defaultExcelTypeFormater = new DefaultExcelTypeFormater();
                }
            }
            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);

            var mainPropertieList = typeof(T).ToColumnDic();

            IList<IExcelExportFormater<ExcelRangeBase>> excelTypes = new List<IExcelExportFormater<ExcelRangeBase>>();
            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = 1;
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
                excelType.SetHeaderCell()?.Invoke(ws1.Cells[row, column], item.Value.Name);
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
                        excelType.SetBodyCell()?.Invoke(ws1.Cells[row, column], mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return ep;

        }

        public static ExcelPackage AddSheet(this ExcelPackage ep, DataTable data)
        {
            ExcelWorkbook wb = ep.Workbook;
            string sheetName = data.TableName;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);

            var headerNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerNames.Add(data.Columns[i].ColumnName);
            }
            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = 1;
            int column = 1;

            //表头行
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(ws1.Cells[row, column], headerName);
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
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(ws1.Cells[row, column], mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return ep;

        }

        public static ExcelPackage AddSheet(this ExcelPackage ep, string sheetName, IList<HeaderInfo> headers, IList<IList<ExportCellValue<ExcelRangeBase>>> data)
        {
            ep.AddSheetHeader(sheetName, headers);
            ExcelWorkbook wb = ep.Workbook;
            ExcelWorksheet ws = wb.Worksheets[sheetName];
            if (data != null && data.Any())
            {
                IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
                int row = ws.Dimension.Rows + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        var mainValue = item.Value;
                        if (item.ExportFormater != null)
                        {
                            item.ExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], mainValue);
                        }
                        else
                        {
                            defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], mainValue);
                        }
                        column++;
                    }

                    row++;
                }
            }
            return ep;

        }

        public static ExcelPackage AddBody(this ExcelPackage ep, DataTable data)
        {
            ExcelWorkbook wb = ep.Workbook;
            string sheetName = data.TableName;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);


            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();

            int row = ws1.Dimension.Rows + 1;

            //数据行 
            if (data != null && data.Rows.Count > 0)
            {
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    int column = 1;
                    for (int j = 0; i < data.Columns.Count; j++)
                    {
                        var mainValue = data.Rows[i][j];
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(ws1.Cells[row, column], mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return ep;

        }

        public static ExcelPackage AddBody(this ExcelPackage ep, string sheetName, IList<IList<ExportCellValue<ExcelRangeBase>>> data)
        { 
            ExcelWorkbook wb = ep.Workbook;
            ExcelWorksheet ws = wb.Worksheets[sheetName];
            if (data != null && data.Any())
            {
                IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
                int row = ws.Dimension.Rows + 1;
                foreach (var dic in data)
                {

                    int column = 1;
                    foreach (var item in dic)
                    {
                        var mainValue = item.Value;
                        if (item.ExportFormater != null)
                        {
                            item.ExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], mainValue);
                        }
                        else
                        {
                            defaultExcelExportFormater.SetBodyCell()?.Invoke(ws.Cells[row, column], mainValue);
                        }
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
            ExcelWorkbook wb = ep.Workbook;
            IExcelTypeFormater<ExcelWorksheet> defaultExcelTypeFormater = new DefaultExcelTypeFormater();

            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);


            IExcelExportFormater<ExcelRangeBase> defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(ws1.Cells[row, column], item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(ws1.Cells[row, column], item.HeaderName);
                }
                column++;
            }

            row++;


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


        public static List<KeyValuePair<PropertyInfo, ExcelColumnAttribute>> ToColumnDic(this Type @type)
        {
            Dictionary<PropertyInfo, ExcelColumnAttribute> mainDic = new Dictionary<PropertyInfo, ExcelColumnAttribute>();
            @type.GetProperties().ToList().ForEach(o =>
            {
                var attribute = o.GetCustomAttribute<ExcelColumnAttribute>();
                if (attribute == null)
                {
                    int order = 1;
                    if (mainDic.Count > 0)
                    {
                        order = mainDic.ElementAt(mainDic.Count - 1).Value.Order + 1;
                    }
                    attribute = new ExcelColumnAttribute(o.Name, order);
                    mainDic.Add(o, attribute);
                }
                else if (!attribute.Ignore)
                {
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
    }
}
