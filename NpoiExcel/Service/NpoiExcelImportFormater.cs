using CExcel.Service;
using System;
using System.Collections.Generic;
using System.Text;

namespace NpoiExcel.Service
{ 
    public class NpoiExcelImportFormater : IExcelImportFormater
    {
        public virtual object Transformation(object origin)
        {
            return origin;
        }
    }
}
