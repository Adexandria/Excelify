using Excelify.Models;
using System.Data;


namespace Excelify.Services
{
    public abstract class ExcelService : IExcelService
    {
        public abstract Task<DataTable> ImportSheet(ISheet sheet);
        public abstract Task<IList<T>> ImportToEntity<T>(ISheet sheet) where T : class;
    }
}
