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
        /// ����
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
                    Name = $"����{i}",
                    Sex = i % 2 == 0 ? 1 : 2,
                    Email = $"coobeedior{i}@123.com",
                    CreateAt = DateTime.Now.AddDays(-1).AddMinutes(i),
                    Email2 = $"���������ʽ"
                };
                students.Add(student);
            }
            try
            {

                var excelPackage = exportService.Export<Student>(students).AddSheet<Student>().AddSheet<Student>().AddSheet<Student>().AddSheet<Student>();

                FileInfo fileInfo = new FileInfo(Path.Combine(path, "���󵼳�.xlsx"));
                excelPackage.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {

            }

        }



        /// <summary>
        /// ����
        /// </summary>
        [TestMethod]
        public void ExportHeader()
        {
            var headers = new List<HeaderInfo>()
            {
               new HeaderInfo("����",(cell,o)=>
               {
                   cell.Value=o;
                   cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                   cell.Style.Fill.BackgroundColor.SetColor(Color.Red);
               } ),
               new HeaderInfo("�Ա�1") ,
               new HeaderInfo("�Ա�2") ,
               new HeaderInfo("�Ա�3") ,
               new HeaderInfo("�Ա�4") ,
               new HeaderInfo("ͷ��") ,

            };

            IList<IList<object>> list = new List<IList<object>>();
            for (int i = 0; i < 10; i++)
            {
                IList<object> cellValues = new List<object>();
                cellValues.Add(new
                {
                    Value = $"����{i}",

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
            FileInfo fileInfo = new FileInfo(Path.Combine(path, "���鵼��.xlsx"));
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
                    Name = $"����{i}",
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
                dc.AutoIncrement = true;//�Զ�����
                dc.AutoIncrementSeed = 1;//��ʼΪ1
                dc.AutoIncrementStep = 1;//����Ϊ1
                dc.AllowDBNull = false;//

                dc = tblDatas.Columns.Add("Product", Type.GetType("System.String"));
                dc = tblDatas.Columns.Add("Version", Type.GetType("System.String"));
                dc = tblDatas.Columns.Add("Description", Type.GetType("System.String"));

                DataRow newRow;
                newRow = tblDatas.NewRow();
                newRow["Product"] = "������";
                newRow["Version"] = "2.0";
                newRow["Description"] = "�Һ�ϲ��";
                tblDatas.Rows.Add(newRow);

                newRow = tblDatas.NewRow();
                newRow["Product"] = "�λ�����";
                newRow["Version"] = "3.0";
                newRow["Description"] = "�ȴ󻰸�����";
                tblDatas.Rows.Add(newRow);
                var excelPackage = workbookBuilder.CreateWorkbook().AddSheet(tblDatas);
                FileInfo fileInfo = new FileInfo(Path.Combine(path, "dataTable����.xlsx"));
                excelPackage.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {

            }

        }

        /// <summary>
        /// ����
        /// </summary>
        [TestMethod]
        public void Import()
        {
            ExcelPackage ep = null;

            try
            {
                using (var fs = File.Open(Path.Combine(path, "���󵼳�.xlsx"), FileMode.Open))
                {
                    ep = workbookBuilder.CreateWorkbook(fs);
                }
                var result = excelImportService.Import<Student>(ep);

            }
            catch (ExportExcelException ex)
            {
                ep.AddErrors<Student>(ex.ExportExcelErrors);
                FileInfo fileInfo = new FileInfo(Path.Combine(path, "���󵼳��쳣.xlsx"));
                ep.SaveAs(fileInfo);
            }
            catch (Exception ex) { }

        }

        /// <summary>
        /// �������
        /// </summary>
        [TestMethod]
        public void AddError()
        {
            try
            {
                var fs = File.Open(Path.Combine(path, "���󵼳�.xlsx"), FileMode.Open);
                var ep = workbookBuilder.CreateWorkbook(fs);
                fs.Close();
                IList<ExportExcelError> errors = new List<ExportExcelError>();
                ExportExcelError a = new ExportExcelError(2, 3, "�����");
                ExportExcelError b = new ExportExcelError(3, 3, "�����11133");
                errors.Add(a);
                errors.Add(b);

                ep.AddErrors<Student>(errors);
                var fs1 = File.Open(Path.Combine(path, "�ֶ������쳣����.xlsx"), FileMode.Open, FileAccess.ReadWrite);
                ep.SaveAs(fs1);
            }
            catch (Exception ex)
            {
            }

        }

        /// <summary>
        /// ����
        /// </summary>
        [TestMethod]
        public void ImportGames()
        {
            ExcelPackage ep = null;

            try
            {
                using (var fs = File.Open("E:\\������Ϸ�˵����� .xlsx", FileMode.Open))
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


    [Excel("��˾", true)]
    public class Game
    {

        [ExcelColumn("��Ϸ���")]
        public string Number { get; set; }
        [ExcelColumn("������")]
        public string AtivationCode { get; set; }


        [ExcelColumn("��Ϸ����")]
        public string Name { get; set; }

        [ExcelColumn("��������")]
        public string TianYi { get; set; }

        [ExcelColumn("�ٶ�����")]
        public string BaiDu { get; set; }

        [ExcelColumn("���")]
        public string Quark { get; set; }

        [ExcelColumn("��Ʒ")]
        public string Gift { get; set; }


        [ExcelColumn("��Ϸ����")]
        public string Description { get; set; }


        [ExcelColumn("��ϷͼƬ1")]
        public string Img1 { get; set; }

        [ExcelColumn("��ϷͼƬ2")]
        public string Img2 { get; set; }

        [ExcelColumn("��ϷͼƬ3")]
        public string Img3 { get; set; }

        [ExcelColumn("��ϷͼƬ4")]
        public string Img4 { get; set; }

        [ExcelColumn("��ϷͼƬ5")]
        public string Img6 { get; set; }

        [ExcelColumn("��Ϸ��Ƶ")]
        public string Video { get; set; }
    }


    [Excel("ѧ����Ϣ", true, typeof(StudentExcelTypeFormater))]
    public class Student
    {
        /// <summary>
        /// ����
        /// </summary>
        //[ExcelColumn("Id", 1)]
        public int Id { get; set; }

        [ExcelColumn("����")]
        public string Name { get; set; }

        /// <summary>
        /// �Ա� ���ӵ���͵�������
        /// </summary>
        [ExcelColumn("�Ա�", 3, typeof(SexExcelTypeFormater), typeof(SexExcelImportFormater))]
        public int Sex { get; set; }

        /// <summary>
        /// ����
        /// </summary>
        [ExcelColumn("����", 4)]
        [EmailAddress]
        public string Email { get; set; }

        /// <summary>
        /// ���� 
        /// </summary>
        [ExcelColumn("����2", 4)]
        [EmailAddress(ErrorMessage="�����ݲ��������ʽ")]
        public string Email2 { get; set; }

        /// <summary>
        /// ����ʱ�� ���˴��ֶ�
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
                c.Value = "���Դ����Դ����Դ����Դ����Դ����Դ����Դ��";
                s.Cells["A1:E1"].Merge = true;//�ϲ���Ԫ��
                s.View.FreezePanes(3, 1); //������
                s.Cells.Style.ShrinkToFit = true;//��Ԫ���Զ���Ӧ��С
                s.Row(1).Height = 44;//�����и�
                s.Row(1).CustomHeight = true;//�Զ������и�
                c.Style.Font.Size = 22;

                #region ���õ�Ԫ����뷽ʽ   
                c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//ˮƽ����
                c.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//��ֱ����
                #endregion

                //c.AutoFitColumns();//��Ԫ��Ŀ��
                c.Worksheet.Cells[c.Worksheet.Dimension.Address].AutoFitColumns();
                c.Worksheet.Cells.AutoFitColumns(2, 50);
                #region ���õ�Ԫ��߿�
                c.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));//���õ�Ԫ�����б߿�
                c.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;//�������õ�Ԫ��ײ��߿���ʽ����ɫ���������Ҿ��ɷֿ����ã�
                c.Style.Border.Bottom.Color.SetColor(Color.LightYellow);
                #endregion


                //var address = typeof(Student).GetCellAddress(nameof(Student.Email));
                //address = $"{address}2:{address}1000";
                //var val2 = s.DataValidations.AddCustomValidation(address);
                //val2.ShowErrorMessage = true;
                //val2.ShowInputMessage = true;
                //val2.PromptTitle = "�Զ��������ϢPromptTitle";
                //val2.Prompt = "�Զ������Prompt";
                //val2.ErrorTitle = "����������ErrorTitle";
                //val2.Error = "����������Error";
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
                        c.Value = "��";
                    }
                    else if (intValue == 2)
                    {
                        c.Value = "Ů";
                    }
                    else
                    {
                        c.Value = "δ֪";
                    }

                }
                else
                {
                    c.Value = "δ֪";
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
            else if (origin?.ToString() == "��")
            {
                return 1;
            }
            else if (origin?.ToString() == "Ů")
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
