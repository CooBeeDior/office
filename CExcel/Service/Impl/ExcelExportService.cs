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
        private readonly IWorkbookBuilder<ExcelPackage> _workbookBuilder;
        public ExcelExportService(IWorkbookBuilder<ExcelPackage> workbookBuilder)
        {
            _workbookBuilder = workbookBuilder;
        }
        public ExcelPackage Export<T>(IList<T> data = null) where T : class, new()
        {
            var ep = _workbookBuilder.CreateWorkbook();
            return ep.AddSheet(data);
        }
        public ExcelPackage Export(IList<object> data)
        {
            var ep = _workbookBuilder.CreateWorkbook();
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
