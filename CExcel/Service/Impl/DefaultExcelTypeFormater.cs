using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
namespace CExcel.Service.Impl
{
    /// <summary>
    /// excel格式化
    /// </summary>
    public class DefaultExcelTypeFormater : IExcelTypeFormater<ExcelWorksheet>
    {
        public virtual Action<ExcelWorksheet> SetExcelWorksheet()
        {
            return (s) =>
            {

                //#region 公式计算  
                //s.Cells["D2:D5"].Formula = "B2*C2";//这是乘法的公式，意思是第二列乘以第三列的值赋值给第四列，这种方法比较简单明了
                //s.Cells[6, 2, 6, 4].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 2, 5, 2).Address);//这是自动求和的方法，至于subtotal的用法你需要自己去了解了
                //#endregion

                //#region 设置sheet背景              
                //s.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //s.Cells.Style.Fill.BackgroundColor.SetColor(Color.LightGray);//设置背景色
                //s.View.ShowGridLines = false;//去掉sheet的网格线
                //s.BackgroundImage.Image = Image.FromFile(@"firstbg.jpg");//设置背景图片
                //#endregion

                //#region  插入图片
                //ExcelPicture picture = s.Drawings.AddPicture("logo", Image.FromFile(@"firstbg.jpg"));//插入图片
                //picture.SetPosition(100, 100);//设置图片的位置
                //picture.SetSize(100, 100);//设置图片的大小
                //#endregion

                //#region 　插入形状
                //ExcelShape shape = s.Drawings.AddShape("shape", eShapeStyle.Rect);//插入形状
                //shape.Font.Color = Color.Red;//设置形状的字体颜色
                //shape.Font.Size = 15;//字体大小
                //shape.Font.Bold = true;//字体粗细
                //shape.Fill.Style = eFillStyle.NoFill;//设置形状的填充样式
                //shape.Border.Fill.Style = eFillStyle.NoFill;//边框样式
                //shape.SetPosition(200, 300);//形状的位置
                //shape.SetSize(80, 30);//形状的大小
                //shape.Text = "test";//形状的内容
                //#endregion

                //#region 超链接
                ////给图片加超链接
                //s.Drawings.AddPicture("logo", Image.FromFile(@"firstbg.jpg"), new ExcelHyperLink("http:\\www.baidu.com", UriKind.Relative));
                ////　 给单元格加超链接
                //s.Cells[1, 1].Hyperlink = new ExcelHyperLink("http:\\www.baidu.com", UriKind.Relative);
                //#endregion

                //#region 隐藏sheet
                //s.Hidden = eWorkSheetHidden.Hidden;//隐藏sheet
                //s.Column(1).Hidden = true;//隐藏某一列
                //s.Row(1).Hidden = true;//隐藏某一行
                //#endregion

                //#region 嵌入VBA代码
                //s.CodeModule.Name = "sheet";
                //s.CodeModule.Code = "code";
                //#endregion

                //#region Excel加密和锁定
                //s.Protection.IsProtected = true;//设置是否进行锁定
                //s.Protection.SetPassword("yk");//设置密码
                //s.Protection.AllowAutoFilter = false;//下面是一些锁定时权限的设置
                //s.Protection.AllowDeleteColumns = false;
                //s.Protection.AllowDeleteRows = false;
                //s.Protection.AllowEditScenarios = false;
                //s.Protection.AllowEditObject = false;
                //s.Protection.AllowFormatCells = false;
                //s.Protection.AllowFormatColumns = false;
                //s.Protection.AllowFormatRows = false;
                //s.Protection.AllowInsertColumns = false;
                //s.Protection.AllowInsertHyperlinks = false;
                //s.Protection.AllowInsertRows = false;
                //s.Protection.AllowPivotTables = false;
                //s.Protection.AllowSelectLockedCells = false;
                //s.Protection.AllowSelectUnlockedCells = false;
                //s.Protection.AllowSort = false;
                //#endregion

                //#region 图表
                //ExcelChart chart = s.Drawings.AddChart("chart", eChartType.ColumnClustered);//eChartType中可以选择图表类型

                //ExcelChartSerie serie = chart.Series.Add(s.Cells[2, 3, 5, 3], s.Cells[2, 1, 5, 1]);//设置图表的x轴和y轴
                //serie.HeaderAddress = s.Cells[1, 3];//设置图表的图例

                //chart.SetPosition(150, 10);//设置位置
                //chart.SetSize(500, 300);//设置大小
                //chart.Title.Text = "销量走势";//设置图表的标题
                //chart.Title.Font.Color = Color.FromArgb(89, 89, 89);//设置标题的颜色
                //chart.Title.Font.Size = 15;//标题的大小
                //chart.Title.Font.Bold = true;//标题的粗体
                //chart.Style = eChartStyle.Style15;//设置图表的样式
                //chart.Legend.Border.LineStyle = eLineStyle.Solid;
                //chart.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);//设置图例的样式
                //#endregion

                //#region 属性设置 针对整个Excel本身的一些其他设置
                //s.Workbook.Properties.Title = "inventory";//设置excel的标题
                //s.Workbook.Properties.Author = "mei";//作者
                //s.Workbook.Properties.Comments = "this is a test";//备注
                //s.Workbook.Properties.Company = "ABC";//公司
                //#endregion

                //#region 下拉框
                //var val = s.DataValidations.AddListValidation("A2:A1000");//设置下拉框显示的数据区域
                //val.Formula.ExcelFormula = "=parameter";//数据区域的名称
                //val.Prompt = "下拉选择参数";//下拉提示
                //val.ShowInputMessage = false;//显示提示内容

                //val.ErrorTitle = "错误标题";

                //val.Formula.Values.Add("aa");
                //val.Formula.Values.Add("bb");
                //val.Formula.Values.Add("cc");
                //#endregion

                //#region 数据校验
                //////时间
                //var val1 = s.DataValidations.AddTimeValidation("A1");
                //// Alternatively:
                //// var validation = sheet.Cells["A1"].DataValidation.AddTimeDataValidation();
                //val1.ShowErrorMessage = true;
                //val1.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                //val1.ShowInputMessage = true;
                //val1.PromptTitle = "Enter time in format HH:MM:SS";
                //val1.Prompt = "Should be greater than 13:30:10";
                //val1.Operator = ExcelDataValidationOperator.greaterThan;
                //var time = val1.Formula.Value;
                //time.Hour = 13;
                //time.Minute = 30;
                //time.Second = 10;


                //var val2 = s.DataValidations.AddCustomValidation("A2:A100");
                //val2.ShowErrorMessage = true;
                //val2.ShowInputMessage = true;
                //val2.PromptTitle = "自定义错误信息PromptTitle";
                //val2.Prompt = "自定义错误Prompt";
                //val2.ErrorTitle = "请输入邮箱ErrorTitle";
                //val2.Error = "请输入邮箱Error";
                //val2.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                //var formula = val2.Formula;
                //formula.ExcelFormula = $"=COUNTIF(A2:A100,\"?*@*.*\")";
                //#endregion
            };
        }


    }
}
