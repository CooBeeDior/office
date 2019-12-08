using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Attributes
{

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExportColumnAttribute : Attribute
    {
        public string Name { get; set; }
        public int Order { get; set; }
        /// <summary>
        ///必须继承 IExcelTypeFormater
        /// </summary>
        public Type ExcelType { get; set; }

        public ExportColumnAttribute(string Name, int Order, Type ExcelType = null)
        {
            this.Name = Name;
            this.Order = Order;
            this.ExcelType = ExcelType;
        }

    }
}
