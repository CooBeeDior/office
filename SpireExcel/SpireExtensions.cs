using CExcel.Service;
using Microsoft.Extensions.DependencyInjection;
using Spire.Xls;
using SpireExcel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Extensions.DependencyInjection
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
