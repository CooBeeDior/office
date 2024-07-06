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



        [Excel("¹«Ë¾", true)]
        public class Game
        {

            [ExcelColumn("ÓÎÏ·±àºÅ")]
            public string Number { get; set; }
            [ExcelColumn("¼¤»îÂë")]
            public string AtivationCode { get; set; }


            [ExcelColumn("ÓÎÏ·Ãû³Æ")]
            public string Name { get; set; }

            [ExcelColumn("ÌìÒíÔÆÅÌ")]
            public string TianYi { get; set; }

            [ExcelColumn("°Ù¶ÈÍøÅÌ")]
            public string BaiDu { get; set; }

            [ExcelColumn("¿ä¿Ë")]
            public string Quark { get; set; }

            [ExcelColumn("ÔùÆ·")]
            public string Gift { get; set; }


            [ExcelColumn("ÓÎÏ·ÃèÊö")]
            public string Description { get; set; }


            [ExcelColumn("ÓÎÏ·Í¼Æ¬1")]
            public string Img1 { get; set; }

            [ExcelColumn("ÓÎÏ·Í¼Æ¬2")]
            public string Img2 { get; set; }

            [ExcelColumn("ÓÎÏ·Í¼Æ¬3")]
            public string Img3 { get; set; }

            [ExcelColumn("ÓÎÏ·Í¼Æ¬4")]
            public string Img4 { get; set; }

            [ExcelColumn("ÓÎÏ·Í¼Æ¬5")]
            public string Img6 { get; set; }

            [ExcelColumn("ÓÎÏ·ÊÓÆµ")]
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
