using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Service
{
    public interface IExcelImportFormater
    {
        object Transformation(object origin);
    }
}
