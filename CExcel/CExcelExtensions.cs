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
            services.AddSingleton<IExcelExportService<ExcelPackage>, ExcelExportService>();
            services.AddSingleton<IExcelImportService<ExcelPackage>, ExcelImportService>();

            services.AddSingleton<IExcelExportFormater, DefaultExcelExportFormater>();
            services.AddSingleton<IExcelTypeFormater, DefaultExcelTypeFormater>();
            services.AddSingleton<IExcelImportFormater, DefaultExcelImportFormater>();
            services.AddSingleton<IExcelProvider<ExcelPackage>, ExcelProvider>(); 
            //services.Add(new ServiceDescriptor(typeof(IExcelProvider<>), typeof(ExcelProvider), ServiceLifetime.Singleton));

            return services;
        }
    }
}
