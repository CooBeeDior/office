using Microsoft.Extensions.DependencyInjection;
using SpireExcel;
using System;
using System.Collections.Generic;
using System.Text;

namespace CExcel.Test
{
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
