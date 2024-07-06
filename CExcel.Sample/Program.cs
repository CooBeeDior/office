using CExcel.Attributes;
using CExcel.Exceptions;
using CExcel.Extensions;
using CExcel.Service;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using Spire.Xls;
using SpireExcel.Extensions;
using System;
using System.Collections.Generic;
using System.IO;

namespace CExcel.Sample
{
    public class Program
    {

        public static void Main(string[] args)
        { 
            CreateHostBuilder(args).Build().Run();
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

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
    }

    public static class Ioc
    {
        private static IServiceCollection service = new ServiceCollection();
        public static IServiceProvider AddCExcelService()
        {
            service.AddCExcelService();
            return service.BuildServiceProvider();
        }

        public static IServiceProvider AddSpireExcelService()
        {
            service.AddSpireExcelService();
            return service.BuildServiceProvider();
        }


        public static IServiceProvider AddNpoiExcelService()
        {
            service.AddNpoiExcelService();
            return service.BuildServiceProvider();
        }
        private static IServiceProvider _provider = null;
        public static IServiceProvider Provider
        {
            get
            {
                if (_provider == null)
                {
                    //lock(obj){}
                    _provider = service.BuildServiceProvider();
                }
                return _provider;
            }
        }
    }
}
