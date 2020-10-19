using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public interface IWorkbookBuilder<TWorkbook>
    {
        TWorkbook CreateWorkbook(); 


        TWorkbook CreateWorkbook(Stream sm); 


        TWorkbook CreateWorkbook(byte[] buffer);



        TWorkbook CreateWorkbook(string filename);

    }
}
