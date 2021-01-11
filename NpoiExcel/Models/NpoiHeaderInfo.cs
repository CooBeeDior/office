using CExcel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NpoiExcel.Models
{ 
    public class NpoiHeaderInfo : HeaderInfo<ICell>
    {
        public NpoiHeaderInfo(string headerName, Action<ICell, object> action = null) : base(headerName, action)
        {

        }
    }
}
