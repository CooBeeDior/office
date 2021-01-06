using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
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
        public static Workbook AddSheet<T>(this Workbook wb, IList<T> data = null) where T : class, new()
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
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater<Worksheet>;
                }
                else
                {
                    defaultExcelTypeFormater = new SpireExcelTypeFormater();
                }
            }
            Worksheet ws1 = wb.Worksheets[sheetName];
            if (ws1 == null)
            {
                ws1 = wb.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);

            var mainPropertieList = typeof(T).ToColumnDic();


            IList<IExcelExportFormater<CellRange>> excelTypes = new List<IExcelExportFormater<CellRange>>();
            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (ws1.CellRecords.LastRow == -1 ? 0 : ws1.CellRecords.LastRow) + 1;
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
                excelType.SetHeaderCell()?.Invoke(ws1[row, column], item.Value.Name);
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
                        excelType.SetBodyCell()?.Invoke(ws1[row, column], mainValue);
                        column++;
                    }
                    row++;
                }
            }
            return wb;
        }
        public static Workbook AddSheet(this Workbook wb, DataTable data)
        {
            string sheetName = data.TableName;
            IExcelTypeFormater<Worksheet> defaultExcelTypeFormater = new SpireExcelTypeFormater();

            Worksheet ws1 = wb.Worksheets[sheetName];
            if (ws1 == null)
            {
                ws1 = wb.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);

            var headerNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                headerNames.Add(data.Columns[i].ColumnName);
            }

            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (ws1.CellRecords.LastRow == -1 ? 0 : ws1.CellRecords.LastRow) + 1;
            int column = 1;

            //表头行
            foreach (var headerName in headerNames)
            {
                defaultExcelExportFormater.SetHeaderCell()?.Invoke(ws1[row, column], headerName);
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
                        defaultExcelExportFormater.SetBodyCell()?.Invoke(ws1[row, column], mainValue);
                        column++;
                    }
                    row++;

                }
            }
            return wb;

        }
        public static Workbook AddSheetHeader(this Workbook wb, string sheetName, IList<SpireHeaderInfo> headers)
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

            Worksheet ws1 = wb.Worksheets[sheetName];
            if (ws1 == null)
            {
                ws1 = wb.Worksheets.Add(sheetName);
            }
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);


            IList<IExcelExportFormater<CellRange>> excelTypes = new List<IExcelExportFormater<CellRange>>();
            IExcelExportFormater<CellRange> defaultExcelExportFormater = new SpireExcelExportFormater();
            int row = (ws1.CellRecords.LastRow == -1 ? 0 : ws1.CellRecords.LastRow) + 1;
            int column = 1;

            //表头行
            foreach (var item in headers)
            {
                if (item.Action == null)
                {
                    defaultExcelExportFormater.SetHeaderCell()(ws1[row, column], item.HeaderName);
                }
                else
                {
                    item.Action.Invoke(ws1[row, column], item.HeaderName);
                }
                column++;
            }

            row++;


            return wb;

        }

        public static Workbook AddErrors(this Workbook wb, string sheetName, IList<ExportExcelError> errors, Action<CellRange, string> action = null)
        {
            if (errors == null || !errors.Any())
            {
                return wb;
            }
            var workSheet = wb.Worksheets[sheetName];
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
            return wb;

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
