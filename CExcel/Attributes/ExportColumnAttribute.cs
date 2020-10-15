using CExcel.Service;
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
        /// 导出Excel，必须继承 IExcelExportFormater,默认：DefaultExcelExportFormater
        /// </summary>
        public Type ExportExcelType { get; set; }

        /// <summary>
        /// 导入Excel,必须继承IExcelImportFormater
        /// </summary>
        public Type ImportExcelType { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Order"></param>
        /// <param name="ExcelType">导出Excel，必须继承 <see cref="IExcelExportFormater">IExcelExportFormater</see>,默认：<see cref="CExcel.Service.Impl.DefaultExcelExportFormater">DefaultExcelExportFormater</see></param>
        /// <param name="ImportExcelType">导入Excel，必须继承<see cref="IExcelImportFormater">IExcelImportFormater</see></param>
        public ExportColumnAttribute(string Name = null, int Order = 0, Type ExportExcelType = null, Type ImportExcelType = null)
        {
            this.Name = Name;
            this.Order = Order;
            if (ExportExcelType != null)
            {
                if (!typeof(IExcelExportFormater).IsAssignableFrom(ExportExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelExportFormater】");
                }

                this.ExportExcelType = ExportExcelType;
            }
            if (ImportExcelType != null)
            {
                if (!typeof(IExcelImportFormater).IsAssignableFrom(ImportExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelImportFormater】");
                }

                this.ImportExcelType = ImportExcelType;
            }

        }

    }
}
