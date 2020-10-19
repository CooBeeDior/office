using CExcel.Service;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Text;

namespace SpireExcel
{
    /// <summary>
    /// Excel格式化
    /// </summary>
    public class SpireExcelTypeFormater : IExcelTypeFormater<Worksheet>
    {
        public virtual Action<Worksheet> SetExcelWorksheet()
        {
            return (s) =>
            {

            };
        }
    }
}
