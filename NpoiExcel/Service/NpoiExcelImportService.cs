using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
using CExcel.Service;
using NPOI.SS.UserModel;
using NpoiExcel.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NpoiExcel.Service
{
    /// <summary>
    /// 导入服务
    /// </summary>
    /// <exception cref="ExportExcelException">导出数据校验不通过</exception>
    public class NpoiExcelImportService : IExcelImportService<IWorkbook>
    {
        public IList<T> Import<T>(IWorkbook workbook, string sheetName = null) where T : class, new()
        {
            ISheet sheet = null;
            if (string.IsNullOrEmpty(sheetName))
            {
                var arrtibute = typeof(T).GetCustomAttribute<ExcelAttribute>();
                if (arrtibute != null)
                {
                    sheet = workbook.GetSheet(arrtibute.SheetName);
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }

            }
            else
            {
                sheet = workbook.GetSheet(sheetName);
            }
            var mainDic = typeof(T).ToColumnDic();
            int totalRows = sheet.LastRowNum + 1;
            int totalColums = sheet.GetRow(0)?.LastCellNum ?? 0;

            IList<T> list = new List<T>();
            //表头行
            int row = 0;
            Dictionary<PropertyInfo, Tuple<ExcelColumnAttribute, IEnumerable<ValidationAttribute>>> filterDic = new Dictionary<PropertyInfo, Tuple<ExcelColumnAttribute, IEnumerable<ValidationAttribute>>>();
            for (int i = 1; i <= totalColums; i++)
            {
                var dic = mainDic.Where(o => o.Value.Name.Equals(sheet.GetRow(row).GetCell(i).ToValue()) || o.Key.Name.Equals(sheet.GetRow(row).GetCell(i).ToValue())).FirstOrDefault();
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
                        //TODO 根据类型获取数据
                        object cellValue = sheet.GetRow(row)?.GetCell(column)?.ToValue();
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
