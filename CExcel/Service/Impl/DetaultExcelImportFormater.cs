using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service.Impl
{
    public class DetaultExcelImportFormater : IExcelImportFormater
    {
        public object Transformation(object origin)
        {
            return origin;
        }
    }
}
