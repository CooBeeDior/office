using CExcel.Service;
using CExcel.Service.Impl;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Extensions.DependencyInjection
{
    public static class CExcelExtensions
    {
        public static IServiceCollection AddCExcelService(this IServiceCollection services)
        {
            //services.AddSingleton<IExcelExportService<ExcelPackage>, ExcelExportService>();
            //services.AddSingleton<IExcelImportService<ExcelPackage>, ExcelImportService>();
            //services.AddSingleton<IExcelProvider<ExcelPackage>, ExcelProvider>();

            services.Add(new ServiceDescriptor(typeof(IExcelProvider<ExcelPackage>), typeof(ExcelProvider), ServiceLifetime.Singleton));

            return services;
        }
    }
}
