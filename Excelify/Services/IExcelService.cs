using Excelify.Models;
using Excelify.Services.Utility;
using NPOI.SS.UserModel;
using System.Data;


namespace Excelify.Services
{
    public interface IExcelService
    {
        IList<T> ImportToEntity<T> (IImportSheet sheet) where T : class;
        IList<T> ImportToEntity<T> (IImportSheet sheet, IExcelMapper excelifyMapper) where T : class;
        ISheet ExportToExcel<T>(IEntityExport<T> dataExport) where T : class;
        DataTable ImportSheet(IImportSheet sheet);
    }
}
