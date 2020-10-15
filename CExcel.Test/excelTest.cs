using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using CExcel.Attributes;
using CExcel.Extensions;
using CExcel.Service;
using CExcel.Service.Impl;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;

namespace CExcel.Test
{
    [TestClass]
    public class excelTest
    {
        /// <summary>
        /// 导出
        /// </summary>
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
                    Email = $"aaa{i}@123.com",
                    //CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                };
                students.Add(student);
            }
            try
            {
                var exportService = new ExcelExportService();

                var excelPackage = exportService.Export<Student>(students).AddSheet<Student>().AddSheet<Student>().AddSheet<Student>().AddSheet<Student>();


                FileInfo fileInfo = new FileInfo("a.xlsx");
                excelPackage.SaveAs(fileInfo);
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


        [ExportColumn("性别", 3, typeof(SexExcelTypeFormater), typeof(SexExcelImportFormater))]
        public int Sex { get; set; }


        [ExportColumn("邮箱", 4)]
        public string Email { get; set; }

        //[ExportColumn("创建时间", 4, typeof(CreateAtExcelTypeFormater), typeof(CreateAtExcelImportFormater))]
        //public DateTime CreateAt { get; set; }
    }
    public class CreateAtExcelImportFormater : DefaultExcelImportFormater
    {
        public override object Transformation(object origin)
        {

            var date = DateTime.ParseExact(origin.ToString(), "yyyy年MM月dd日 HH:mm:ss", null);
            return date;
        }
    }

    public class CreateAtExcelTypeFormater : DefaultExcelExportFormater
    {
        public override Action<ExcelRangeBase, object> SetBodyCell()
        {
            return (c, o) =>
            {
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.Green);
                c.Style.Numberformat.Format = "yyyy年MM月dd日 HH:mm:ss"; 
                c.Style.ShrinkToFit = false;//单元格自动适应大小
                //c.AddComment(o.ToString(), $"时间:{o.ToString("yyyy/MM/dd HH:mm:ss")}");
           
                //c.Worksheet.Column(typeof(Student).GetPropertyIndex(nameof(Student.CreateAt))).Width = 50;
                c.Value = o;
            };
        }

        public override Action<ExcelRangeBase, object> SetHeaderCell()
        {

            return (c, o) =>
            {
                base.SetHeaderCell()(c, o);
                c.Style.Font.Color.SetColor(Color.Black);//字体颜色
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.Red);
                c.AddComment(o?.ToString() ?? "", "超级管理员1");

            };
        }
    }


    public class StudentExcelTypeFormater : DefaultExcelTypeFormater
    {
        public override Action<ExcelWorksheet> SetExcelWorksheet()
        {
            return (s) =>
            {
                var address = typeof(Student).GetPropertyAddress(nameof(Student.Email));
                address = $"{address}2:{address}1000";
                var val2 = s.DataValidations.AddCustomValidation(address);
                val2.ShowErrorMessage = true;
                val2.ShowInputMessage = true;
                val2.PromptTitle = "自定义错误信息PromptTitle";
                val2.Prompt = "自定义错误Prompt";
                val2.ErrorTitle = "请输入邮箱ErrorTitle";
                val2.Error = "请输入邮箱Error";
                val2.ErrorStyle = ExcelDataValidationWarningStyle.stop;                
                var formula = val2.Formula;
                formula.ExcelFormula = $"=COUNTIF({address},\"?*@*.*\")";
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
