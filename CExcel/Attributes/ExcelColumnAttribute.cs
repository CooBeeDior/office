using CExcel.Service;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Attributes
{

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// 排序
        /// </summary>
        public int Order { get; set; }
        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool Ignore { get; }
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
        /// <param name="name"></param>
        /// <param name="order"></param>
        /// <param name="ignore"></param>
        /// <param name="exportExcelType">导出Excel，必须继承 <see cref="IExcelExportFormater">IExcelExportFormater</see>,默认：<see cref="CExcel.Service.Impl.DefaultExcelExportFormater">DefaultExcelExportFormater</see></param>
        /// <param name="importExcelType">导入Excel，必须继承<see cref="IExcelImportFormater">IExcelImportFormater</see></param>
        public ExcelColumnAttribute(string name, int order, bool ignore = false, Type exportExcelType = null, Type importExcelType = null)
        {
            this.Name = name;
            this.Order = order;
            this.Ignore = ignore;
            if (exportExcelType != null)
            {
                var type = typeof(IExcelExportFormater<>).MakeGenericType(exportExcelType.GetInterfaces()[0].GenericTypeArguments[0]);
                if (!type.IsAssignableFrom(exportExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelExportFormater】");
                }

                this.ExportExcelType = exportExcelType;
            }
            if (importExcelType != null)
            {
                if (!typeof(IExcelImportFormater).IsAssignableFrom(importExcelType))
                {
                    throw new ArgumentException("not assignablefrom 【IExcelImportFormater】");
                }

                this.ImportExcelType = importExcelType;
            }

        }


        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="name"></param>
        /// <param name="order"></param> 
        /// <param name="exportExcelType">导出Excel，必须继承 <see cref="IExcelExportFormater">IExcelExportFormater</see>,默认：<see cref="CExcel.Service.Impl.DefaultExcelExportFormater">DefaultExcelExportFormater</see></param>
        /// <param name="importExcelType">导入Excel，必须继承<see cref="IExcelImportFormater">IExcelImportFormater</see></param>
        public ExcelColumnAttribute(string name, int order, Type exportExcelType, Type importExcelType) : this(name, order, false, exportExcelType, importExcelType)
        {

        }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="name"></param> 
        /// <param name="exportExcelType">导出Excel，必须继承 <see cref="IExcelExportFormater">IExcelExportFormater</see>,默认：<see cref="CExcel.Service.Impl.DefaultExcelExportFormater">DefaultExcelExportFormater</see></param>
        /// <param name="importExcelType">导入Excel，必须继承<see cref="IExcelImportFormater">IExcelImportFormater</see></param>
        public ExcelColumnAttribute(string name, Type exportExcelType, Type importExcelType) : this(name, 0, false, exportExcelType, importExcelType)
        {

        }
        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="name"></param>
        /// <param name="order"></param> 
        public ExcelColumnAttribute(string name, int order) : this(name, order, false, null, null)
        {

        }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="name"></param>
        public ExcelColumnAttribute(string name) : this(name, 0, false, null, null)
        {

        }

    }
}
