using Excelify.Models;
using Excelify.Services.Utility;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;


namespace Excelify.Services
{
    public abstract class ExcelService : IExcelService
    {
        public abstract Stream ExportToExcel<T>(IEntityExport<T> dataExport, string fileName) where T : class;
        public abstract XSSFWorkbook ExportToExcel<T>(IEntityExport<T> dataExport) where T : class;
        public abstract DataTable ImportSheet(IImportSheet sheet);
        public abstract IList<T> ImportToEntity<T>(IImportSheet sheet) where T : class;
        public abstract IList<T> ImportToEntity<T>(IImportSheet sheet, IExcelMapper excelifyMapper) where T : class;
        public abstract void SetSheetName(int sheetName);
    }
}
