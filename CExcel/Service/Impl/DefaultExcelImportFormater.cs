using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service.Impl
{
    /// <summary>
    /// 导入格式化
    /// </summary>
    public class DefaultExcelImportFormater : IExcelImportFormater
    {
        public virtual object Transformation(object origin)
        {
            return origin;
        }
    }
}
