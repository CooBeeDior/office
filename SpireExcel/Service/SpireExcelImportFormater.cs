using CExcel.Service;
using System;
using System.Collections.Generic;
using System.Text;

namespace SpireExcel
{
    public class SpireExcelImportFormater : IExcelImportFormater
    {
        public virtual object Transformation(object origin)
        {
            return origin;
        }
    }
}
