using Excelify.Models;
using Excelify.Services.Utility;
using NPOI.SS.UserModel;
using System.Data;


namespace Excelify.Services
{
    public abstract class ExcelService : IExcelService
    {
        public abstract ISheet ExportToExcel<T>(IEntityExport<T> dataExport) where T : class;
        public abstract DataTable ImportSheet(IImportSheet sheet);
        public abstract IList<T> ImportToEntity<T>(IImportSheet sheet) where T : class;
        public abstract IList<T> ImportToEntity<T>(IImportSheet sheet, IExcelMapper excelifyMapper) where T : class;
    }
}
