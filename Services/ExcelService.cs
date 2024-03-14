using Excelify.Models;
using System.Data;


namespace Excelify.Services
{
    public abstract class ExcelService : IExcelService
    {
        public abstract Task<DataTable> ImportSheet(IExcelifySheet sheet);
        public abstract Task<IList<T>> ImportToEntity<T>(IExcelifySheet sheet) where T : class;
    }
}
