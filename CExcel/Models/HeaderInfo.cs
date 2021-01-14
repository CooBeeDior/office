using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace CExcel
{
    [DebuggerDisplay("名称: {HeaderName}")]
    public class HeaderInfo<TExcelRange>
    {
        public HeaderInfo(string headerName, Action<TExcelRange, object> action = null)
        {
            if (string.IsNullOrEmpty(headerName))
            {
                throw new ArgumentNullException(nameof(headerName));
            }
            this.HeaderName = headerName;
            this.Action = action;
        }
        public string HeaderName { get; }

        public Action<TExcelRange, object> Action { get; }
    }

    public class HeaderInfo : HeaderInfo<ExcelRangeBase>
    {
        public HeaderInfo(string headerName, Action<ExcelRangeBase, object> action = null) : base(headerName, action)
        {

        }
    }
}
