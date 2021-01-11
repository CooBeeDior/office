using CExcel.Service;
using NPOI.SS.UserModel;
using NpoiExcel.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NpoiExcel.Service
{ 
    /// <summary>
    /// 导出服务
    /// </summary>
    public class NpoiExcelExportService : IExcelExportService<IWorkbook>
    {
        private readonly IWorkbookBuilder<IWorkbook> _workbookBuilder;
        public NpoiExcelExportService(IWorkbookBuilder<IWorkbook> workbookBuilder)
        {
            _workbookBuilder = workbookBuilder;
        }
        public IWorkbook Export<T>(IList<T> data = null) where T : class, new()
        {
            var workbook = _workbookBuilder.CreateWorkbook();
            return workbook.AddSheet(data);
        }
        public IWorkbook Export(IList<object> data)
        {
            var workbook = _workbookBuilder.CreateWorkbook();
            return workbook.AddSheet(data);
        }


        public Stream ToStream(IWorkbook workbook)
        {
            MemoryStream sm = new MemoryStream();
            workbook.Write(sm);
            return sm;
        }
    }
}
