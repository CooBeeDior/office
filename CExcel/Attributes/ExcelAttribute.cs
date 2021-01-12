using CExcel.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExcelAttribute : Attribute
    {
        public string SheetName { get; }

        public bool IsIncrease { get; }
        /// <summary>
        ///导出Excel，必须继承 IExcelExportFormater,默认：DefaultExcelExportFormater
        /// </summary>
        public Type ExportExcelType { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="IsIncrease"></param>
        /// <param name="ExcelType">必须继承 <see cref="CExcel.Service.IExcelTypeFormater">IExcelTypeFormater</see>,默认：<see cref="CExcel.Service.Impl.DefaultExcelTypeFormater">DefaultExcelTypeFormater</see> </param>
        public ExcelAttribute(string sheetName, bool isIncrease = false, Type exportExcelType = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            this.SheetName = sheetName;
            this.IsIncrease = isIncrease;
            if (exportExcelType != null)
            {
                var genericType = exportExcelType.GetInterfaces()?.FirstOrDefault()?.GenericTypeArguments?.FirstOrDefault();
                if (genericType == null)
                {
                    throw new ArgumentException("not assignablefrom 【IExcelTypeFormater】");
                }
                var type = typeof(IExcelTypeFormater<>).MakeGenericType(genericType);
                if (!type.IsAssignableFrom(exportExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelTypeFormater】");
                }

                this.ExportExcelType = exportExcelType;
            }

        }
    }
}
