using CExcel.Attributes;
using CExcel.Extensions;
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
    /// 导出服务
    /// </summary>
    public class ExcelExportService : IExcelExportService<ExcelPackage>
    {
        public ExcelPackage Export<T>(IList<T> data = null) where T : class, new()
        {
            ExcelPackage ep = new ExcelPackage();
            return ep.AddSheet(data);
        }

        public Stream ToStream(ExcelPackage workbook)
        {
            MemoryStream sm = new MemoryStream();
            workbook.SaveAs(sm);
            return sm;
        }
    }
}
