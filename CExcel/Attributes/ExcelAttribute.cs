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
        ///必须继承 IExcelTypeFormater
        /// </summary>
        public Type ExcelType { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="IsIncrease"></param>
        /// <param name="ExcelType">必须继承 IExcelTypeFormater</param>
        public ExcelAttribute(string SheetName = null, bool IsIncrease = false, Type ExcelType = null)
        {
            this.SheetName = SheetName;
            this.IsIncrease = IsIncrease;
            if (ExcelType != null)
            {
                if (typeof(IExcelTypeFormater).IsAssignableFrom(ExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelTypeFormater】");
                }

                this.ExcelType = ExcelType;
            }

        }
    }
}
