using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Service;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Spire.Xls;
using SpireExcel;
using SpireExcel.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace CExcel.Test
{
    [TestClass]
    public class spireExcelTest
    {
        private readonly IExcelExportService<Workbook> exportService = null;
        private readonly IWorkbookBuilder<Workbook> workbookBuilder = null;
        private readonly IExcelImportService<Workbook> excelImportService = null;
        public spireExcelTest()
        {
            var provider = Ioc.AddSpireExcelService();
            workbookBuilder = provider.GetService<IWorkbookBuilder<Workbook>>();
            excelImportService = provider.GetService<IExcelImportService<Workbook>>();
            exportService = provider.GetService<IExcelExportService<Workbook>>();
        }
        /// <summary>
        /// 导出
        /// </summary>
        [TestMethod]
        public void Export()
        {
            var aa = workbookBuilder.CreateWorkbook();

            IList<Student> students = new List<Student>();
            for (int i = 0; i < 100; i++)
            {
                Student student = new Student()
                {
                    Id = i,
                    Name = $"姓名{i}",
                    Sex = 2,
                    Email = $"aaa{i}@123.com",
                    //CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                };
                students.Add(student);
            }
            try
            {
                var excelPackage = exportService.Export<Student>(students).AddSheet<Student>(students).AddSheet<Student>().AddSheet<Student>().AddSheet<Student>();


                excelPackage.SaveToFile("spirea.xlsx", FileFormat.Version2016);
                excelPackage.SaveToFile("spirea.pdf", FileFormat.PDF);
            }
            catch (Exception ex)
            {

            }

        }

        /// <summary>
        /// 导入
        /// </summary>
        [TestMethod]
        public void Import()
        {
            Workbook wb = null;
            try
            {
                using (var fs = File.Open("spirea.xlsx", FileMode.Open))
                    wb = workbookBuilder.CreateWorkbook(fs);

                var result = excelImportService.Import<Student>(wb);

            }
            catch (ExportExcelException ex)
            {
                wb.AddErrors<Student>(ex.ExportExcelErrors);
                wb.SaveToFile("spireb.xlsx");
            }
            catch (Exception ex) { }

        }


        [Excel("学生信息")]
        public class Student
        {
            /// <summary>
            /// 主键
            /// </summary>
            [ExcelColumn("Id", 1)]
            public int Id { get; set; }

            //[ExcelColumn("姓名", 2)]
            //[EmailAddress(ErrorMessage = "不是邮箱格式")]
            public string Name { get; set; }


            [ExcelColumn("性别", 3, typeof(SexExcelTypeFormater), typeof(SexExcelImportFormater))]
            [EmailAddress(ErrorMessage = "不是邮箱格式")]
            public int Sex { get; set; }


            //[ExcelColumn("邮箱", 4)]
            [EmailAddress]
            public string Email { get; set; }

            [IngoreExcelColumn]
            public DateTime CreateAt { get; set; }
        }





        public class SexExcelTypeFormater : SpireExcelExportFormater
        {
            public override Action<CellRange, object> SetBodyCell()
            {
                return (c, o) =>
                {
                    base.SetBodyCell()(c, o);
                    if (int.TryParse(o.ToString(), out int intValue))
                    {
                        if (intValue == 1)
                        {
                            c.Value = "男";
                        }
                        else if (intValue == 2)
                        {
                            c.Value = "女";
                        }
                        else
                        {
                            c.Value = "未知";
                        }

                    }
                    else
                    {
                        c.Value = "未知";
                    }

                };
            }


        }

        public class SexExcelImportFormater : IExcelImportFormater
        {
            public object Transformation(object origin)
            {
                if (origin == null)
                {
                    return 0;
                }
                else if (origin?.ToString() == "男")
                {
                    return 1;
                }
                else if (origin?.ToString() == "女")
                {
                    return 2;
                }
                else
                {
                    return 0;
                }
            }
        }
    }
}
