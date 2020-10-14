using CExcel.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// 导入
    /// </summary>
    public class ExcelImportService : IExcelImportService<ExcelPackage>
    {
        public IList<T> Import<T>(ExcelPackage workbook, string sheetName) where T : class, new()
        {
            ExcelWorksheet ws1 = null;
            if (string.IsNullOrEmpty(sheetName))
            {
                ws1 = workbook.Workbook.Worksheets[1];
            }
            else
            {
                ws1 = workbook.Workbook.Worksheets[sheetName];
            }


            Dictionary<PropertyInfo, ExportColumnAttribute> mainDic = new Dictionary<PropertyInfo, ExportColumnAttribute>();

            typeof(T).GetProperties().ToList().ForEach(o =>
            {
                var attribute = o.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    mainDic.Add(o, attribute);
                }
            });
            //var mainPropertieList = mainDic.OrderBy(o => o.Value.Order).ToList();

            int totalRows = ws1.Dimension.Rows;
            int totalColums = ws1.Dimension.Columns;

            IList<T> list = new List<T>();
            //表头行
            int row = 1;
            Dictionary<PropertyInfo, ExportColumnAttribute> filterDic = new Dictionary<PropertyInfo, ExportColumnAttribute>();
            for (int i = 1; i <= totalColums; i++)
            {
                var dic = mainDic.Where(o => o.Value.Name.Equals(ws1.Cells[row, i].Value?.ToString()?.Trim()) || o.Key.Name.Equals(ws1.Cells[row, i].Value?.ToString()?.Trim())).FirstOrDefault();
                if (dic.Key != null)
                {
                    filterDic.Add(dic.Key, dic.Value);
                }
               
            }

            row++;

            IList<IExcelImportFormater> excelTypes = new List<IExcelImportFormater>();

            for (int i = row; i <= totalRows; i++)
            {
                T t = new T();
                int column = 1;
                foreach (var item in filterDic)
                {
                    var property = item.Key;
                    if (property != null)
                    {
                        object cellValue = ws1.GetValue(row, column);
                        if (item.Value.ImportExcelType != null)
                        {
                            var excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.ImportExcelType.FullName).FirstOrDefault();
                            if (excelType == null)
                            {
                                excelType = Activator.CreateInstance(item.Value?.ImportExcelType) as IExcelImportFormater;
                                excelTypes.Add(excelType);
                            }
                            cellValue = excelType.Transformation(cellValue);
                        }                  
                        if (cellValue == null)
                        {
                            cellValue = "";
                        }
                        else if (property.PropertyType == typeof(string))
                        {
                            cellValue = cellValue.ToString();
                        }
                        else if (property.PropertyType == typeof(int))
                        {
                            cellValue = Convert.ToInt32(cellValue);
                        }
                        else if (property.PropertyType == typeof(long))
                        {
                            cellValue = Convert.ToInt64(cellValue);
                        }
                        else if (property.PropertyType == typeof(double))
                        {
                            cellValue = Convert.ToDecimal(cellValue);
                        }
                        else if (property.PropertyType == typeof(decimal))
                        {
                            cellValue = Convert.ToDecimal(cellValue);
                        }
                        else if (property.PropertyType == typeof(DateTime))
                        {
                            cellValue = Convert.ToDateTime(cellValue);

                        }
                        else
                        {
                            cellValue = cellValue.ToString();
                        }
                        property?.SetValue(t, cellValue);
                    }


                    column++;
                }
                list.Add(t);
            }
            return list;
        }
    }
}
