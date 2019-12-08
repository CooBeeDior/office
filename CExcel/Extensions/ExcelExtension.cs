using CExcel.Attributes;
using CExcel.Service;
using CExcel.Service.Impl;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CExcel.Extensions
{
    public static class ExcelExtension
    {
        public static ExcelPackage AddSheet<T>(this ExcelPackage ep, IList<T> list) where T : class, new()
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
                    sheetName = $"{excelAttribute.SheetName}{wb.Worksheets.Count + 1}";
                }
                else
                {
                    sheetName = excelAttribute.SheetName;
                }
                if (excelAttribute.ExcelType != null)
                {
                    defaultExcelTypeFormater = Activator.CreateInstance(excelAttribute.ExcelType) as IExcelTypeFormater;
                }
                else
                {
                    defaultExcelTypeFormater = new DefaultExcelTypeFormater();
                }
            }

            ExcelWorksheet ws1 = wb.Worksheets.Add(sheetName);
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


            IList<IExcelTypeFormater> excelTypes = new List<IExcelTypeFormater>();
            int row = 1;
            int column = 1;

            //表头行
            foreach (var item in mainPropertieList)
            {
                IExcelTypeFormater excelType = null;
                if (item.Value.ExcelType != null)
                {
                    excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ExcelType.FullName).FirstOrDefault();
                    if (excelType == null)
                    {
                        excelType = Activator.CreateInstance(item.Value.ExcelType) as IExcelTypeFormater;
                        excelTypes.Add(excelType);
                    }
                }
                else
                {
                    excelType = defaultExcelTypeFormater;
                }
                excelType.SetHeaderCell()?.Invoke(ws1.Cells[row, column], item.Value.Name);
                column++;
            }

            row++;

            //数据行 
            foreach (var item in list)
            {
                column = 1;
                foreach (var mainPropertie in mainPropertieList)
                {
                    IExcelTypeFormater excelType = null;
                    var mainValue = mainPropertie.Key.GetValue(item);
                    if (mainPropertie.Value.ExcelType != null)
                    {
                        excelType = excelTypes.Where(o => o.GetType().FullName == mainPropertie.Value.ExcelType.FullName).FirstOrDefault();
                        if (excelType == null)
                        {
                            excelType = Activator.CreateInstance(mainPropertie.Value.ExcelType) as IExcelTypeFormater;
                            excelTypes.Add(excelType);
                        }
                    }
                    else
                    {
                        excelType = defaultExcelTypeFormater;
                    }
                    excelType.SetBodyCell()?.Invoke(ws1.Cells[row, column], mainValue);
                    column++;
                }
                row++;
            }

            return ep;

        }
    }
}
