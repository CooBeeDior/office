using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Service;
using Spire.Xls;
using SpireExcel.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SpireExcel
{
    public class SpireExcelImportService : IExcelImportService<Workbook>
    {
        public IList<T> Import<T>(Workbook workbook, string sheetName = null) where T : class, new()
        {
            Worksheet ws1 = null;
            if (string.IsNullOrEmpty(sheetName))
            {
                var arrtibute = typeof(T).GetCustomAttribute<ExcelAttribute>();
                if (arrtibute != null)
                {
                    ws1 = workbook.Worksheets[arrtibute.SheetName];
                }
                else
                {
                    ws1 = workbook.Worksheets[1];
                }

            }
            else
            {
                ws1 = workbook.Worksheets[sheetName];
            }

            Dictionary<PropertyInfo, ExcelColumnAttribute> mainDic = new Dictionary<PropertyInfo, ExcelColumnAttribute>();

            typeof(T).GetProperties().ToList().ForEach(o =>
            {
                var attribute = o.GetCustomAttribute<ExcelColumnAttribute>();
                if (attribute != null)
                {
                    mainDic.Add(o, attribute);
                }
            });
            //var mainPropertieList = mainDic.OrderBy(o => o.Value.Order).ToList();

            int totalRows = ws1.Rows.Count();
            int totalColums = ws1.Columns.Count();

            IList<T> list = new List<T>();
            //表头行
            int row = 1;
            Dictionary<PropertyInfo, Tuple<ExcelColumnAttribute, IEnumerable<ValidationAttribute>>> filterDic = new Dictionary<PropertyInfo, Tuple<ExcelColumnAttribute, IEnumerable<ValidationAttribute>>>();
            for (int i = 1; i <= totalColums; i++)
            {
                var dic = mainDic.Where(o => o.Value.Name.Equals(ws1[row, i].Value2?.ToString()?.Trim()) || o.Key.Name.Equals(ws1[row, i].Value2?.ToString()?.Trim())).FirstOrDefault();
                if (dic.Key != null)
                {
                    var validationAttributes = dic.Key.GetCustomAttributes<ValidationAttribute>();
                    filterDic.Add(dic.Key, Tuple.Create(dic.Value, validationAttributes));
                }

            }

            row++;

            IList<IExcelImportFormater> excelTypes = new List<IExcelImportFormater>();
            IList<ExportExcelError> errors = new List<ExportExcelError>();
            bool flag = true;
            for (int i = row; i <= totalRows; i++)
            {
                T t = new T();
                int column = 1;


                foreach (var item in filterDic)
                {
                    var property = item.Key;
                    if (property != null)
                    {
                        object cellValue = ws1[row,column].Value;
                        if (item.Value.Item2 != null && item.Value.Item2.Any())
                        {
                            foreach (var validator in item.Value.Item2)
                            {
                                if (!validator.IsValid(cellValue))
                                {
                                    errors.Add(new ExportExcelError(row, column, validator.ErrorMessage));
                                    flag = false;
                                }
                            }
                        }
                        if (flag)
                        {
                            if (item.Value.Item1.ImportExcelType != null)
                            {
                                var excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.Item1.ImportExcelType.FullName).FirstOrDefault();
                                if (excelType == null)
                                {
                                    excelType = Activator.CreateInstance(item.Value.Item1.ImportExcelType) as IExcelImportFormater;
                                    excelTypes.Add(excelType);
                                }
                                cellValue = excelType.Transformation(cellValue);
                            }
                            if (cellValue == null)
                            {
                                cellValue = null;
                            }
                            else if (property.PropertyType == typeof(string))
                            {
                                cellValue = cellValue.ToString();
                            }
                            else if (property.PropertyType == typeof(char) || property.PropertyType == typeof(char?))
                            {
                                cellValue = Convert.ToChar(cellValue);
                            }
                            else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                            {
                                cellValue = Convert.ToInt32(cellValue);
                            }
                            else if (property.PropertyType == typeof(long) || property.PropertyType == typeof(long?))
                            {
                                cellValue = Convert.ToInt64(cellValue);
                            }
                            else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(double?))
                            {
                                cellValue = Convert.ToDecimal(cellValue);
                            }
                            else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                            {
                                cellValue = Convert.ToDecimal(cellValue);
                            }
                            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                            {
                                cellValue = Convert.ToDateTime(cellValue);

                            }
                            else
                            {
                                cellValue = cellValue.ToString();
                            }
                            property?.SetValue(t, cellValue);
                        }
                    }


                    column++;
                }
                row++;
                list.Add(t);
            }
            if (!flag)
            {
                throw new ExportExcelException(errors);
            }
            return list;
        }



       
    }
}
