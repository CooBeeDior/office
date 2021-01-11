using CExcel.Models;
using CExcel.Service;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NpoiExcel.Service
{
    public class NpoiWorkbookBuilder : IWorkbookBuilder<IWorkbook>
    {

        public IWorkbook CreateWorkbook(CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            IWorkbook workbook = null;
            if (excelVersion == CExcelVersion.Version2003)
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                workbook= new XSSFWorkbook();
            } 
            return workbook;

        }

        public IWorkbook CreateWorkbook(Stream sm, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            IWorkbook workbook = null;
            if (excelVersion == CExcelVersion.Version2003)
            {
                workbook = new HSSFWorkbook(sm);
            }
            else
            {
                workbook = new XSSFWorkbook(sm);
            }
            return workbook;
        }

        public IWorkbook CreateWorkbook(byte[] buffer, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            IWorkbook workbook = null;
            if (excelVersion == CExcelVersion.Version2003)
            {
                workbook = new HSSFWorkbook(new MemoryStream( buffer));
            }
            else
            {
                workbook = new XSSFWorkbook(new MemoryStream(buffer));
            }
            return workbook;
        }

        public IWorkbook CreateWorkbook(string filename, CExcelVersion excelVersion = CExcelVersion.Version2007)
        {
            IWorkbook workbook = null;
            if (excelVersion == CExcelVersion.Version2003)
            {
                var fs = File.OpenRead(filename);
                workbook = new HSSFWorkbook(fs);
            }
            else
            {
                workbook = new XSSFWorkbook(filename);
            }
            return workbook;
        }
    }
}
