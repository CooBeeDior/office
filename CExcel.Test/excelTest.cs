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
using System.Reflection;
using System.Data;
using CExcel.Models;
using NPOI.SS.UserModel;

namespace CExcel.Test
{
    [TestClass]
    public class excelTest
    {
        private readonly IExcelExportService<ExcelPackage> exportService = null;
        private readonly IExcelImportService<ExcelPackage> excelImportService = null;
        private readonly IWorkbookBuilder<ExcelPackage> workbookBuilder;
        private readonly string path = "excel";
        public excelTest()
        {
            var provider = Ioc.AddCExcelService();
            exportService = provider.GetService<IExcelExportService<ExcelPackage>>();
            excelImportService = provider.GetService<IExcelImportService<ExcelPackage>>();
            workbookBuilder = provider.GetService<IWorkbookBuilder<ExcelPackage>>();
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

        }



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
                    Sex = i % 2 == 0 ? 1 : 2,
                    Email = $"coobeedior{i}@123.com",
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                    Email2 = $"不是邮箱格式"
                };
                students.Add(student);
            }
            try
            {

                var excelPackage = exportService.Export<Student>(students).AddSheet<Student>().AddSheet<Student>().AddSheet<Student>().AddSheet<Student>();

                FileInfo fileInfo = new FileInfo(Path.Combine(path, "对象导出.xlsx"));
                excelPackage.SaveAs(fileInfo);
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
            var headers = new List<HeaderInfo>()
            {
               new HeaderInfo("姓名",(cell,o)=>
               {
                   cell.Value=o;
                   cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                   cell.Style.Fill.BackgroundColor.SetColor(Color.Red);
               } ),
               new HeaderInfo("性别1") ,
               new HeaderInfo("性别2") ,
               new HeaderInfo("性别3") ,
               new HeaderInfo("性别4") ,
               new HeaderInfo("头像") ,

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
                    ExportFormater = new SexExcelTypeFormater()
                });
                cellValues.Add(new
                {
                    Value = i % 3,
                    ExportFormater = new SexExcelTypeFormater()
                });
                cellValues.Add(new
                {
                    Value = i % 3,
                    ExportFormater = new SexExcelTypeFormater()
                });

                cellValues.Add(new
                {
                    Value = $"http://www.baidu.com/{i}",
                    aa = new ImageExcelTypeFormater()
                });
                list.Add(cellValues);

            }

            var ep = workbookBuilder.CreateWorkbook().AddSheetHeader("cc", headers).AddBody("cc", list);
            FileInfo fileInfo = new FileInfo(Path.Combine(path, "数组导出.xlsx"));
            ep.SaveAs(fileInfo);
        }

        [TestMethod]
        public void ExportFromDatatable()
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
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                };
                students.Add(student);
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
                newRow["Product"] = "大话西游";
                newRow["Version"] = "2.0";
                newRow["Description"] = "我很喜欢";
                tblDatas.Rows.Add(newRow);

                newRow = tblDatas.NewRow();
                newRow["Product"] = "梦幻西游";
                newRow["Version"] = "3.0";
                newRow["Description"] = "比大话更幼稚";
                tblDatas.Rows.Add(newRow);
                var excelPackage = workbookBuilder.CreateWorkbook().AddSheet(tblDatas);
                FileInfo fileInfo = new FileInfo(Path.Combine(path, "dataTable导出.xlsx"));
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
            ExcelPackage ep = null;

            try
            {
                using (var fs = File.Open(Path.Combine(path, "对象导出.xlsx"), FileMode.Open))
                {
                    ep = workbookBuilder.CreateWorkbook(fs);
                }
                var result = excelImportService.Import<Student>(ep);

            }
            catch (ExportExcelException ex)
            {
                ep.AddErrors<Student>(ex.ExportExcelErrors);
                FileInfo fileInfo = new FileInfo(Path.Combine(path, "对象导出异常.xlsx"));
                ep.SaveAs(fileInfo);
            }
            catch (Exception ex) { }

        }

        /// <summary>
        /// 导入错误
        /// </summary>
        [TestMethod]
        public void AddError()
        {
            try
            {
                var fs = File.Open(Path.Combine(path, "对象导出.xlsx"), FileMode.Open);
                var ep = workbookBuilder.CreateWorkbook(fs);
                fs.Close();
                IList<ExportExcelError> errors = new List<ExportExcelError>();
                ExportExcelError a = new ExportExcelError(2, 3, "错误的");
                ExportExcelError b = new ExportExcelError(3, 3, "错误的11133");
                errors.Add(a);
                errors.Add(b);

                ep.AddErrors<Student>(errors);
                var fs1 = File.Open(Path.Combine(path, "手动增加异常导出.xlsx"), FileMode.Open, FileAccess.ReadWrite);
                ep.SaveAs(fs1);
            }
            catch (Exception ex)
            {
            }

        }

        /// <summary>
        /// 导入
        /// </summary>
        [TestMethod]
        public void ImportGames()
        {
            ExcelPackage ep = null;

            try
            {
                using (var fs = File.Open("E:\\最新游戏菜单下载 .xlsx", FileMode.Open))
                {
                    ep = workbookBuilder.CreateWorkbook(fs);
                }
                var result = excelImportService.Import<Game>(ep);

            }
            catch (ExportExcelException ex)
            {
                ep.AddErrors<Student>(ex.ExportExcelErrors);
                FileInfo fileInfo = new FileInfo("b.xlsx");
                ep.SaveAs(fileInfo);
            }
            catch (Exception ex) { }

        }


    }


    [Excel("公司", true)]
    public class Game
    {

        [ExcelColumn("游戏编号")]
        public string Number { get; set; }
        [ExcelColumn("激活码")]
        public string AtivationCode { get; set; }


        [ExcelColumn("游戏名称")]
        public string Name { get; set; }

        [ExcelColumn("天翼云盘")]
        public string TianYi { get; set; }

        [ExcelColumn("百度网盘")]
        public string BaiDu { get; set; }

        [ExcelColumn("夸克")]
        public string Quark { get; set; }

        [ExcelColumn("赠品")]
        public string Gift { get; set; }


        [ExcelColumn("游戏描述")]
        public string Description { get; set; }


        [ExcelColumn("游戏图片1")]
        public string Img1 { get; set; }

        [ExcelColumn("游戏图片2")]
        public string Img2 { get; set; }

        [ExcelColumn("游戏图片3")]
        public string Img3 { get; set; }

        [ExcelColumn("游戏图片4")]
        public string Img4 { get; set; }

        [ExcelColumn("游戏图片5")]
        public string Img6 { get; set; }

        [ExcelColumn("游戏视频")]
        public string Video { get; set; }
    }


    [Excel("学生信息", true, typeof(StudentExcelTypeFormater))]
    public class Student
    {
        /// <summary>
        /// 主键
        /// </summary>
        //[ExcelColumn("Id", 1)]
        public int Id { get; set; }

        [ExcelColumn("姓名")]
        public string Name { get; set; }

        /// <summary>
        /// 性别 增加导入和导出处理
        /// </summary>
        [ExcelColumn("性别", 3, typeof(SexExcelTypeFormater), typeof(SexExcelImportFormater))]
        public int Sex { get; set; }

        /// <summary>
        /// 邮箱
        /// </summary>
        [ExcelColumn("邮箱", 4)]
        [EmailAddress]
        public string Email { get; set; }

        /// <summary>
        /// 邮箱 
        /// </summary>
        [ExcelColumn("邮箱2", 4)]
        [EmailAddress(ErrorMessage="该数据不是邮箱格式")]
        public string Email2 { get; set; }

        /// <summary>
        /// 创建时间 过滤此字段
        /// </summary>
        [IngoreExcelColumn]
        public DateTime CreateAt { get; set; }
    }





    public class StudentExcelTypeFormater : DefaultExcelTypeFormater
    {
        public override Action<ExcelWorksheet> SetExcelWorksheet()
        {
            return (s) =>
            {
                base.SetExcelWorksheet()(s);

                int row = (s?.Dimension?.Rows ?? 0) + 1;
                int column = 1;
                var c = s.Cells[row, column];
                c.Value = "测试大大测试大大测试大大测试大大测试大大测试大大测试大大";
                s.Cells["A1:E1"].Merge = true;//合并单元格
                s.View.FreezePanes(3, 1); //冻结行
                s.Cells.Style.ShrinkToFit = true;//单元格自动适应大小
                s.Row(1).Height = 44;//设置行高
                s.Row(1).CustomHeight = true;//自动调整行高
                c.Style.Font.Size = 22;

                #region 设置单元格对齐方式   
                c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                c.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                #endregion

                //c.AutoFitColumns();//单元格的宽度
                c.Worksheet.Cells[c.Worksheet.Dimension.Address].AutoFitColumns();
                c.Worksheet.Cells.AutoFitColumns(2, 50);
                #region 设置单元格边框
                c.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));//设置单元格所有边框
                c.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;//单独设置单元格底部边框样式和颜色（上下左右均可分开设置）
                c.Style.Border.Bottom.Color.SetColor(Color.LightYellow);
                #endregion


                //var address = typeof(Student).GetCellAddress(nameof(Student.Email));
                //address = $"{address}2:{address}1000";
                //var val2 = s.DataValidations.AddCustomValidation(address);
                //val2.ShowErrorMessage = true;
                //val2.ShowInputMessage = true;
                //val2.PromptTitle = "自定义错误信息PromptTitle";
                //val2.Prompt = "自定义错误Prompt";
                //val2.ErrorTitle = "请输入邮箱ErrorTitle";
                //val2.Error = "请输入邮箱Error";
                //val2.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                //var formula = val2.Formula;
                //formula.ExcelFormula = $"=COUNTIF({address},\"?*@*.*\")";
            };

        }

    }


    public class SexExcelTypeFormater : DefaultExcelExportFormater
    {
        public override Action<ExcelRangeBase, object> SetBodyCell()
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


    public class ImageExcelTypeFormater : DefaultExcelExportFormater
    {
        public override Action<ExcelRangeBase, object> SetBodyCell()
        {
            return (c, o) =>
            {
                c.Style.Font.Size = 12;
                c.Style.Font.UnderLine = true;
                c.Style.Font.Color.SetColor(Color.Blue);
                c.Hyperlink = new Uri(o.ToString(), UriKind.Absolute);
                c.Value = o;


                var fs = File.OpenRead(@"images/a.jpg");
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
                fs.Close();
                fs.Dispose();
                c.Worksheet.AddPicture(buffer, c, true);


            };
        }


    }
}
