using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
using CExcel.Service;
using CExcel.Service.Impl;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using Microsoft.Extensions.DependencyInjection;
using System.Data;
using NPOI.SS.UserModel;
using NpoiExcel.Service;
using NpoiExcel.Extensions;
using NpoiExcel.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

namespace CExcel.Test
{
    [TestClass]
    public class npoiTest
    {
        private readonly IExcelExportService<IWorkbook> exportService = null;
        private readonly IExcelImportService<IWorkbook> excelImportService = null;
        private readonly IWorkbookBuilder<IWorkbook> workbookBuilder;
        public npoiTest()
        {
            var provider = Ioc.AddNpoiExcelService();
            exportService = provider.GetService<IExcelExportService<IWorkbook>>();
            excelImportService = provider.GetService<IExcelImportService<IWorkbook>>();
            workbookBuilder = provider.GetService<IWorkbookBuilder<IWorkbook>>();
        }



        /// <summary>
        /// 导出
        /// </summary>
        [TestMethod]
        public void Export()
        {
            IList<Student1> Student1s = new List<Student1>();
            for (int i = 0; i < 100; i++)
            {
                Student1 Student1 = new Student1()
                {
                    Id = i,
                    Name = $"姓名{i}",
                    Sex = 2,
                    Email = $"aaa{i}@123.com",
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                    Image = $"图片{i}"
                };
                Student1s.Add(Student1);
            }
            try
            {
                var workBook = workbookBuilder.CreateWorkbook();

                var excelPackage = exportService.Export<Student1>(Student1s).AddSheet<Student1>().AddSheet<Student1>().AddSheet<Student1>().AddSheet<Student1>();

                FileStream fs = File.Create("a.xlsx");
                excelPackage.Write(fs);
                fs.Close();

            }
            catch (Exception ex)
            {

            }

        }



        /// <summary>
        /// 导出
        /// </summary>
        [TestMethod]
        public void ExportHeader()
        {
            var headers = new List<NpoiHeaderInfo>()
            {
               new NpoiHeaderInfo("姓名",(cell,o)=>
               { cell.SetCellValue(o?.ToString());
               var  cellStyle=   cell.Sheet.Workbook.CreateCellStyle();
             cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.FillBackgroundColor=IndexedColors.Red.Index;
               cell.CellStyle=cellStyle;
               } ),
                         new NpoiHeaderInfo("性别") ,                         new NpoiHeaderInfo("性别") ,
                new NpoiHeaderInfo("性别") ,
                                   new NpoiHeaderInfo("头像") ,

            };
            IList<IList<object>> list = new List<IList<object>>();
            for (int i = 0; i < 10; i++)
            {
                IList<object> cellValues = new List<object>();
                cellValues.Add(new
                {
                    Value = $"姓名{i}",

                });

                cellValues.Add(new
                {
                    Value = i % 3,
                    ExportFormater = new Sex1ExcelTypeFormater()
                });
                cellValues.Add(new
                {
                    Value = i % 3,
                    ExportFormater = new Sex1ExcelTypeFormater()
                });
                cellValues.Add(new
                {
                    Value = i % 3,
                    ExportFormater = new Sex1ExcelTypeFormater()
                });

                cellValues.Add(new
                {
                    Value = $"http://www.baidu888.com/{i}",
                    //aa = new Image1ExcelTypeFormater()
                });
                list.Add(cellValues);

            }

            var ep = workbookBuilder.CreateWorkbook().AddSheetHeader("cc", headers).AddBody("cc", list);
            FileStream fs = File.Create("d.xlsx");
            ep.Write(fs);
        }

        [TestMethod]
        public void ExportFromDatatable()
        {

            IList<Student1> Student1s = new List<Student1>();
            for (int i = 0; i < 100; i++)
            {
                Student1 Student1 = new Student1()
                {
                    Id = i,
                    Name = $"姓名{i}",
                    Sex = 2,
                    Email = $"aaa{i}@123.com",
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                };
                Student1s.Add(Student1);
            }
            try
            {

                DataTable tblDatas = new DataTable("Datas");
                DataColumn dc = null;
                dc = tblDatas.Columns.Add("ID", Type.GetType("System.Int32"));
                dc.AutoIncrement = true;//自动增加
                dc.AutoIncrementSeed = 1;//起始为1
                dc.AutoIncrementStep = 1;//步长为1
                dc.AllowDBNull = false;//

                dc = tblDatas.Columns.Add("Product", Type.GetType("System.String"));
                dc = tblDatas.Columns.Add("Version", Type.GetType("System.String"));
                dc = tblDatas.Columns.Add("Description", Type.GetType("System.String"));

                DataRow newRow;
                newRow = tblDatas.NewRow();
                newRow["Product"] = "大话西756游";
                newRow["Version"] = "2.750";
                newRow["Description"] = "我很756喜欢";
                tblDatas.Rows.Add(newRow);

                newRow = tblDatas.NewRow();
                newRow["Product"] = "梦幻756西游";
                newRow["Version"] = "3.0";
                newRow["Description"] = "比大话67更幼稚";
                tblDatas.Rows.Add(newRow);
                var excelPackage = workbookBuilder.CreateWorkbook().AddSheet(tblDatas);
                FileStream fs = File.Create("c.xlsx");
                excelPackage.Write(fs);
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
            using (var fs = File.Open("a.xlsx", FileMode.Open))
            {
                IWorkbook ep = null;
                try
                {

                    ep = workbookBuilder.CreateWorkbook(fs);

                    var result = excelImportService.Import<Student1>(ep); 

                }
                catch (ExportExcelException ex)
                {
                    ep.AddErrors<Student1>(ex.ExportExcelErrors);
                    FileStream fs1 = File.Create("b.xlsx");
                    ep.Write(fs1);
                    fs1.Close();
                }
                catch (Exception ex) { }
            }
        }


        /// <summary>
        /// 导入错误
        /// </summary>
        [TestMethod]
        public void AddError()
        {
            try
            {
                var fs = File.Open("a.xlsx", FileMode.Open);
                var ep = workbookBuilder.CreateWorkbook(fs);
                fs.Close();
                IList<ExportExcelError> errors = new List<ExportExcelError>();
                ExportExcelError a = new ExportExcelError(2, 3, "错误的");
                ExportExcelError b = new ExportExcelError(3, 3, "错误的11133");
                errors.Add(a);
                errors.Add(b);

                ep.AddErrors<Student1>(errors);

                ep.Write(fs);
            }
            catch (Exception ex)
            {
            }

        }


    }

    [Excel("学生信息", true)]
    public class Student1
    {
        /// <summary>
        /// 主键
        /// </summary>
        //[ExcelColumn("Id", 1)]
        public int Id { get; set; }

        [ExcelColumn("姓名")]
        [EmailAddress(ErrorMessage = "不是邮箱格式")]
        public string Name { get; set; }


        [ExcelColumn("性别", 3, typeof(Sex1ExcelTypeFormater), typeof(Sex1ExcelImportFormater))]
        public int Sex { get; set; }


        [ExcelColumn("邮箱", 4)]
        [EmailAddress]
        public string Email { get; set; }

        [ExcelColumn("创建时间", 4)]
        //[IngoreExcelColumn]
        public DateTime CreateAt { get; set; }

        /// <summary>
        /// 图片
        /// </summary>
        [ExcelColumn("图片", 5, typeof(Image1ExcelTypeFormater), null)]
        public string Image { get; set; }
    }





    public class Student1ExcelTypeFormater : NpoiExcelTypeFormater
    {
        public override Action<ISheet> SetExcelWorksheet()
        {
            return (s) =>
            {
                base.SetExcelWorksheet()(s);

                var address = typeof(Student1).GetCellAddress(nameof(Student1.Email));
                address = $"{address}2:{address}1000";

                XSSFDataValidationHelper helper = new XSSFDataValidationHelper((XSSFSheet)s);

                //创建验证规则
                IDataValidationConstraint constraint = helper.CreateCustomConstraint($"=COUNTIF({address},\"?*@*.*\")");

                var validation = helper.CreateValidation(constraint, new CellRangeAddressList(1, 1000, 0, 0));

                //设置约束提示信息
                validation.CreateErrorBox("错误", "请按右侧下拉箭头选择!");
                validation.ShowErrorBox = true;
                validation.ShowPromptBox = true;
                validation.CreateErrorBox("请输入邮箱ErrorTitle", "请输入邮箱Error");
                validation.CreatePromptBox("自定义错误信息PromptTitle", "自定义错误Prompt");
                validation.ErrorStyle = 1;

                s.AddValidationData(validation);


            };

        }

    }



    public class Sex1ExcelTypeFormater : NpoiExcelExportFormater
    {
        public override Action<ICell, object> SetBodyCell()
        {
            return (c, o) =>
            {
                base.SetBodyCell()(c, o);
                if (int.TryParse(o.ToString(), out int intValue))
                {
                    if (intValue == 1)
                    {
                        c.SetCellValue("男");
                    }
                    else if (intValue == 2)
                    {
                        c.SetCellValue("女");

                    }
                    else
                    {
                        c.SetCellValue("未知");
                    }

                }
                else
                {
                    c.SetCellValue("未知");
                }

            };
        }


    }
    public class Sex1ExcelImportFormater : IExcelImportFormater
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


    public class Image1ExcelTypeFormater : NpoiExcelExportFormater
    {
        public override Action<ICell, object> SetBodyCell()
        {
            return (c, o) =>
            {
                var fs = File.OpenRead(@"images/a.jpg");
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
                fs.Close();
                fs.Dispose();
                c.AddPicture(buffer);
            };
        }


    }
}
