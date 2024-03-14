using Excelify.Models;
using System.Data;


namespace Excelify.Services
{
    internal interface IExcelService
    {
        Task<IList<T>> ImportToEntity<T> (ISheet sheet) where T : class;
        Task<DataTable> ImportSheet(ISheet sheet);
    }
}
