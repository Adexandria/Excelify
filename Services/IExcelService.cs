using Excelify.Models;
using System.Data;


namespace Excelify.Services
{
    internal interface IExcelService
    {
        Task<IList<T>> ImportToEntity<T> (IExcelifySheet sheet) where T : class;
        Task<DataTable> ImportSheet(IExcelifySheet sheet);
    }
}
