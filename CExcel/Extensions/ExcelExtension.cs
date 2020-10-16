using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Service;
using CExcel.Service.Impl;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
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
            IExcelTypeFormater defaultExcelTypeFormater = null;
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
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExportExcelType) as IExcelTypeFormater;
                }
                else
                {
                    defaultExcelTypeFormater = new DefaultExcelTypeFormater();
                }
            }
            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
            defaultExcelTypeFormater.SetExcelWorksheet()?.Invoke(ws1);

            Dictionary<PropertyInfo, ExportColumnAttribute> mainDic = new Dictionary<PropertyInfo, ExportColumnAttribute>();

            typeof(T).GetProperties().ToList().ForEach(o =>
            {
                var attribute = o.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    mainDic.Add(o, attribute);
                }
            });
            var mainPropertieList = mainDic.OrderBy(o => o.Value.Order).ToList();


            IList<IExcelExportFormater> excelTypes = new List<IExcelExportFormater>();
            IExcelExportFormater defaultExcelExportFormater = new DefaultExcelExportFormater();
            int row = 1;
            int column = 1;

            //表头行
            foreach (var item in mainPropertieList)
            {
                IExcelExportFormater excelType = null;
                if (item.Value.ExportExcelType != null)
                {
                    excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ExportExcelType.FullName).FirstOrDefault();
                    if (excelType == null)
                    {
                        excelType = Activator.CreateInstance(item.Value.ExportExcelType) as IExcelExportFormater;
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
                        IExcelExportFormater excelType = null;
                        var mainValue = mainPropertie.Key.GetValue(item);
                        if (mainPropertie.Value.ExportExcelType != null)
                        {
                            excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(mainPropertie.Value.ExportExcelType) as IExcelExportFormater;
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
      
        
        /// <summary>
        /// 根据属性名获取所在行地址
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string GetPropertyAddress(this Type obj, string propertyName)
        {
            int index = obj.GetPropertyIndex(propertyName);
            string address = getAddress(index);
            return address;
        }

        /// <summary>
        /// 根据属性名获取该列所在行的索引
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static int GetPropertyIndex(this Type obj, string propertyName)
        {
            var excelAttribute = obj.GetCustomAttribute<ExcelAttribute>();
            if (excelAttribute == null)
            {
                throw new Exception($"类型必须包含{nameof(ExcelAttribute)}特性");
            }


            var properties = obj.GetProperties().Where(p => p.GetCustomAttribute<ExportColumnAttribute>() != null).OrderBy(p => p.GetCustomAttribute<ExportColumnAttribute>().Order).ToList();
            if (!properties.Any(o => o.Name == propertyName))
            {
                throw new Exception($"不存在属性名{propertyName}且定义{nameof(ExportColumnAttribute)}");
            }
            int index = properties.IndexOf(properties.FirstOrDefault(o => o.Name == propertyName));

            return index + 1;
        }
        private static string getAddress(int index)
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
