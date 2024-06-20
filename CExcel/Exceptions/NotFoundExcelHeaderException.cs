using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Exceptions
{
    public class NotFoundExcelHeaderException : Exception
    {
      

        public NotFoundExcelHeaderException(string message="未找到excel表头信息"):base(message)
        {
            
        }
    }
}
