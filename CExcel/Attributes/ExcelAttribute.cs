using CExcel.Service;
using System;
using System.Collections.Generic;
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
        /// <param name="ExcelType">必须继承 IExcelExportFormater,默认：DefaultExcelExportFormater </param>
        public ExcelAttribute(string SheetName = null, bool IsIncrease = false, Type ExportExcelType = null)
        {
            this.SheetName = SheetName;
            this.IsIncrease = IsIncrease;
            if (ExportExcelType != null)
            {
                if (!typeof(IExcelExportFormater).IsAssignableFrom(ExportExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelTypeFormater】");
                }

                this.ExportExcelType = ExportExcelType;
            }

        }
    }
}
