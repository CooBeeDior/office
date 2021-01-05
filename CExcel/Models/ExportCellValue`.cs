using CExcel.Service;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Models
{
    public class ExportCellValue<TExcelRange>
    {
        public object Value { get; set; }

        public IExcelExportFormater<TExcelRange> ExportFormater { get; set; }
    }
}
