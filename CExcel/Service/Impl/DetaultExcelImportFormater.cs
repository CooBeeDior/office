using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// 导入格式化
    /// </summary>
    public class DetaultExcelImportFormater : IExcelImportFormater
    {
        public object Transformation(object origin)
        {
            return origin;
        }
    }
}
