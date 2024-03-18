using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Excelify.Services
{
    public class ExcelifyFactory
    {
        public ExcelifyFactory(Assembly assembly = null)
        {
            var entryTypes = assembly?.GetTypes().Where(s => !s.IsAbstract
           && s.BaseType == typeof(ExcelService));
            if (entryTypes != null)
            {
                excelTypes.AddRange(entryTypes);
            }
            else
            {
                excelTypes.AddRange(Assembly.GetExecutingAssembly().GetTypes().Where(s => !s.IsAbstract
               && s.BaseType == typeof(ExcelService)));
            }
            
          
        }
        public IExcelService CreateService(string extensionType, int _sheetName)
        {
            IExcelService excelService;
            try
            {
                if(extensionType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                || extensionType.Equals("application/vnd.ms-excel"))
                {
                    var excelType = excelTypes.FirstOrDefault() ?? throw new Exception("Excel service does not exist");

                    excelService = Activator.CreateInstance(excelType, _sheetName) as IExcelService;

                }
                else
                {
                    throw new Exception("Invalid extension type", new Exception("Only xlsx and xls types are accepted"));
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Failed to create service", ex);
            }

            return excelService;
        }

        private readonly List<Type> excelTypes = [];
    }
}
