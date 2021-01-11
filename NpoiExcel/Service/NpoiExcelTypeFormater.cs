using CExcel.Service;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NpoiExcel.Service
{
    public class NpoiExcelTypeFormater  :IExcelTypeFormater<ISheet>
    {
        public virtual Action<ISheet> SetExcelWorksheet()
        {
            return (s) =>
            {
                s.CreateFreezePane(1, 2);//冻结行
           
            };
        }
    }
}
