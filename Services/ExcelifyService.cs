using Excelify.Models;
using System.Data;

namespace Excelify.Services
{
    public class ExcelifyService : ExcelService
    {
        public ExcelifyService()
        {
            _sheetName = 0;
        }
        public ExcelifyService(int sheetName)
        {

            _sheetName = sheetName;

        }
        public override Task<DataTable> ImportSheet(ISheet sheet)
        {
            throw new NotImplementedException();
        }

        public override Task<IList<T>> ImportToEntity<T>(ISheet sheet)
        {
            throw new NotImplementedException();
        }

        private readonly int _sheetName;
    }
}
