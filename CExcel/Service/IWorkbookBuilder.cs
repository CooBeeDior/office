using CExcel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CExcel.Service
{
    public interface IWorkbookBuilder<TWorkbook>
    {
        TWorkbook CreateWorkbook(CExcelVersion excelVersion= CExcelVersion.Version2007); 


        TWorkbook CreateWorkbook(Stream sm, CExcelVersion excelVersion = CExcelVersion.Version2007); 


        TWorkbook CreateWorkbook(byte[] buffer, CExcelVersion excelVersion = CExcelVersion.Version2007);



        TWorkbook CreateWorkbook(string filename, CExcelVersion excelVersion = CExcelVersion.Version2007);

    }
}
