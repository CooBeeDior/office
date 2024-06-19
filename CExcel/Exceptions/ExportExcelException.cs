using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Exceptions
{
    public class ExportExcelException : Exception
    {
        public ExportExcelException(IList<ExportExcelError> errors = null)
        {
            ExportExcelErrors = errors ?? new List<ExportExcelError>();

        }
        public IList<ExportExcelError> ExportExcelErrors { get; }


        public override string Message
        {
            get
            {
                StringBuilder stringBuilder = new StringBuilder();
                foreach (var item in ExportExcelErrors)
                {
                    stringBuilder.AppendLine($"第{item.Column}列第{item.Row}行{item.Message};");
                }
                return stringBuilder.ToString();
            }
        }
    }

    public class ExportExcelError
    {
        public ExportExcelError(int row, int column, string message)
        {
            this.Row = row;
            this.Column = column;
            this.Message = message;
        }
        public int Row { get; set; }

        public int Column { get; set; }

        public string Message { get; set; }
    }
}
