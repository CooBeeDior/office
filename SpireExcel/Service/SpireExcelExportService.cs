using CExcel.Service;
using Spire.Xls;
using SpireExcel.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SpireExcel
{
    public class SpireExcelExportService : IExcelExportService<Workbook>
    {
        private readonly IWorkbookBuilder<Workbook> _workbookBuilder; 
        public SpireExcelExportService(IWorkbookBuilder<Workbook> workbookBuilder)
        {
            _workbookBuilder = workbookBuilder;
        }
        public Workbook Export<T>(IList<T> data = null) where T : class, new()
        {
            var workbook = _workbookBuilder.CreateWorkbook();
           
            return workbook.AddSheet<T>(data); 
        }

        public Workbook Export(IList<object> data)
        {
            var workbook = _workbookBuilder.CreateWorkbook();

            return workbook.AddSheet(data);
        }

        public Stream ToStream(Workbook workbook)
        {
            MemoryStream sm = new MemoryStream();
            workbook.SaveToStream(sm);
            return sm;
        }
    }
}
