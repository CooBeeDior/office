using CExcel.Service;
using Microsoft.Extensions.DependencyInjection;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Text;

namespace SpireExcel
{
    public static class SpireExtensions
    {
        public static IServiceCollection AddSpireExcelService(this IServiceCollection services)
        {
            services.AddSingleton<IExcelExportService<Workbook>, SpireExcelExportService>();
            services.AddSingleton<IExcelImportService<Workbook>, SpireExcelImportService>(); 
            services.AddSingleton<IExcelProvider<Workbook>, SpireExcelProvider>();
            services.AddSingleton<IWorkbookBuilder<Workbook>, SpireWorkbookBuilder>();

            return services;
        }
    }
}
