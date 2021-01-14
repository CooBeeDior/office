using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NpoiExcel.Models;

namespace CExcel.Sample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            IList<NpoiHeaderInfo> list = new List<NpoiHeaderInfo>();
            for (int i = 0; i < 100; i++)
            {

                NpoiHeaderInfo headerInfo = new NpoiHeaderInfo(i.ToString());
                list.Add(headerInfo);
            }




            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
    }
}
