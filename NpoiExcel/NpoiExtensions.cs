using CExcel.Service;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NpoiExcel.Service;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Extensions.DependencyInjection
{
     
    public static class NpoiExtensions
    {
        public static IServiceCollection AddNpoiExcelService(this IServiceCollection services)
        {
            services.AddSingleton<IExcelExportService<IWorkbook>, NpoiExcelExportService>();
            services.AddSingleton<IExcelImportService<IWorkbook>, NpoiExcelImportService>();

            services.AddSingleton<IExcelProvider<IWorkbook>, NpoiExcelProvider>();
            services.AddSingleton<IWorkbookBuilder<IWorkbook>, NpoiWorkbookBuilder>();

            return services;
        }
    }
}
