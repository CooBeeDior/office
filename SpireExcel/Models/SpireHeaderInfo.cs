using CExcel;
using Spire.Xls;
using System;

namespace SpireExcel
{
    public class SpireHeaderInfo : HeaderInfo<CellRange>
    {
        public SpireHeaderInfo(string headerName, Action<CellRange, object> action = null) : base(headerName, action)
        {

        }
    }
}
