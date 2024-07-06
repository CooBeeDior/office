using CExcel.Attributes;
using CExcel.Config;
using CExcel.Exceptions;
using CExcel.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// 导入服务
    /// </summary>
    /// <exception cref="ExportExcelException">导出数据校验不通过</exception>
    public class ExcelImportService : IExcelImportService<ExcelPackage>
    {
        private readonly ExcelConfig _excelConfig;
        public ExcelImportService(ExcelConfig excelConfig)
        {
            _excelConfig = excelConfig;
        }
        public IList<T> Import<T>(ExcelPackage workbook, string sheetName = null) where T : class, new()
        {
            ExcelWorksheet sheet = null;
            if (string.IsNullOrEmpty(sheetName))
            {
                var arrtibute = typeof(T).GetCustomAttribute<ExcelAttribute>();
                if (arrtibute != null)
                {
                    sheet = workbook.Workbook.Worksheets[arrtibute.SheetName];
                }
                else
                {
                    sheet = workbook.Workbook.Worksheets[1];
                }

            }
            else
            {
                sheet = workbook.Workbook.Worksheets[sheetName];
            }
            var mainDic = typeof(T).ToColumnDic();
            int totalRows = sheet.Dimension.Rows;
            int totalColums = sheet.Dimension.Columns;

            IList<T> list = new List<T>();
            //表头行
            int row = 1;
            Dictionary<PropertyInfo, Tuple<int, ExcelColumnAttribute, IEnumerable<ValidationAttribute>>> filterDic = new Dictionary<PropertyInfo, Tuple<int, ExcelColumnAttribute, IEnumerable<ValidationAttribute>>>();

            while (row <= _excelConfig.MaxNumberRowsMatchHeader)
            {
                for (int i = 1; i <= totalColums; i++)
                {
                    var dic = mainDic.Where(o => o.Value.Name.Equals(sheet.Cells[row, i].Value?.ToString()?.Trim()) || o.Key.Name.Equals(sheet.Cells[row, i].Value?.ToString()?.Trim())).FirstOrDefault();
                    if (dic.Key != null)
                    {
                        var validationAttributes = dic.Key.GetCustomAttributes<ValidationAttribute>();
                        filterDic.Add(dic.Key, Tuple.Create(i, dic.Value, validationAttributes));
                    }

                }
                row++;
                if (filterDic != null && filterDic.Count > 0)
                {
                    break;
                }
            }
            if (filterDic == null || filterDic.Count == 0)
            {
                throw new NotFoundExcelHeaderException();
            }

            IList<IExcelImportFormater> excelTypes = new List<IExcelImportFormater>();
            IList<ExportExcelError> errors = new List<ExportExcelError>();
            bool flag = true;
            for (int i = row; i <= totalRows; i++)
            {
                T t = new T();
                //int column = 1;


                foreach (var item in filterDic)
                {
                    int column = item.Value.Item1;
                    var property = item.Key;
                    if (property != null)
                    {
                        object cellValue = sheet.GetValue(row, column);
                        if (item.Value.Item2 != null && item.Value.Item3.Any())
                        {
                            foreach (var validator in item.Value.Item3)
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
                            if (item.Value.Item2.ImportExcelType != null)
                            {
                                var excelType = excelTypes.Where(o => o.GetType().FullName == item.Value.Item2.ImportExcelType.FullName).FirstOrDefault();
                                if (excelType == null)
                                {
                                    excelType = Activator.CreateInstance(item.Value.Item2.ImportExcelType) as IExcelImportFormater;
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


                    //column++;
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
