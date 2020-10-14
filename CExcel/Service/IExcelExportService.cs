using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelExportService<TWorkbook>
    {
        TWorkbook Export<T>(IList<T> data = null) where T : class, new();


        Stream ToStream(TWorkbook workbook);
    }
}
