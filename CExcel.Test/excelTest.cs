using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using CExcel.Attributes;
using CExcel.Service;
using CExcel.Service.Impl;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CExcel.Test
{
    [TestClass]
    public class excelTest
    {
        [TestMethod]
        public void Import()
        {
            try
            {
                var excelImportService = new ExcelImportService();
                var fs = File.Open("a.xlsx", FileMode.Open);
                var ep = ExcelExcelPackageBuilder.CreateExcelPackage(fs);

                var result = excelImportService.Import<Student>(ep, "学生信息1");
                fs.Close();
            }
            catch (Exception ex)
            {
            }

        }

        [TestMethod]
        public void Export()
        {

            IList<Student> students = new List<Student>();
            for (int i = 0; i < 100; i++)
            {
                Student student = new Student()
                {
                    Id = i,
                    Name = $"姓名{i}",
                    Sex = 2,
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                };
                students.Add(student);
            }
            try
            {
                var exportService = new ExcelExportService();

                var excelPackage = exportService.Export(students);
                FileInfo fileInfo = new FileInfo("a.xlsx");
                excelPackage.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {

            }

        }
    }


    [Excel("学生信息", true, typeof(StudentExcelTypeFormater))]
    public class Student
    {
        /// <summary>
        /// 主键
        /// </summary>
        [ExportColumn("Id", 1)]
        public int Id { get; set; }

        [ExportColumn("姓名", 2)]
        public string Name { get; set; }


        [ExportColumn("性别", 3, typeof(SexExcelTypeFormater),typeof(SexExcelImportFormater))]
        public int Sex { get; set; }

        [ExportColumn("创建时间", 4)]
        public DateTime CreateAt { get; set; }
    }


    public class StudentExcelTypeFormater : DefaultExcelExportFormater
    {
        public override Action<ExcelRangeBase, object> SetBodyCell()
        {
            return (c, o) =>
            {
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.Green);
                c.AddComment(o.ToString(), "用户1");
                c.Value = o;
            };
        }

        public override Action<ExcelRangeBase, object> SetHeaderCell()
        {
            return (c, o) =>
            {
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.LawnGreen);
                c.Value = o;
            };
        }
    }


    public class SexExcelTypeFormater : DefaultExcelExportFormater
    {
        public override Action<ExcelRangeBase, object> SetBodyCell()
        {
            return (c, o) =>
            {
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
