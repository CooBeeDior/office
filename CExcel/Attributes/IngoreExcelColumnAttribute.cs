using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Attributes
{
    /// <summary>
    /// 忽略字段
    /// </summary>
    public class IngoreExcelColumnAttribute : ExcelColumnAttribute
    {
        public IngoreExcelColumnAttribute() : base(null, 0, true)
        {
        }
    }
}
