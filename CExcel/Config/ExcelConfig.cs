using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Config
{
    public class ExcelConfig
    {
        public ExcelConfig(int maxNumberRowsMatchHeader = 8)
        {
            this.MaxNumberRowsMatchHeader = maxNumberRowsMatchHeader;
        }


        /// <summary>
        /// 最多匹配excel表头多少行
        /// </summary>
        public int MaxNumberRowsMatchHeader { get; }  
    }
}
