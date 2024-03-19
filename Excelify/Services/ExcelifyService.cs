using Excelify.Models;
using Excelify.Services.Extensions;
using Excelify.Services.Utility;
using Excelify.Services.Utility.Attributes;
using NPOI.XSSF.UserModel;
using System.Data;

namespace Excelify.Services
{
    public class ExcelifyService : ExcelService
    {

        public override void SetSheetName(int sheetName = 0)
        {
            _sheetName = sheetName;
        }


        public override DataTable ImportSheet(IImportSheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet), "sheet can not be null");

            return sheet.ExtractValues(_sheetName);
        }

        public override IList<T> ImportToEntity<T>(IImportSheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet), "sheet can not be null");

            var extractedValues = sheet.ExtractValues(_sheetName);
            var entities = _excelifyMapper.Map<T>(extractedValues.Rows.OfType<DataRow>()).Result;
            return entities;
        }

        public override IList<T> ImportToEntity<T>(IImportSheet sheet, IExcelMapper excelifyMapper)
        {
            if (excelifyMapper == null)
                throw new ArgumentNullException(nameof(excelifyMapper), "Excel mapper can not be null");

            var extractedValues = sheet.ExtractValues(_sheetName);
            var entities = excelifyMapper.Map<T>(extractedValues.Rows.OfType<DataRow>()).Result;
            return entities;
        }

        public override Stream ExportToExcel<T>(IEntityExport<T> dataExport, string fileName)
        {
            var extractedAttributes = ExcelifyRecord.GetAttributeProperty<ExcelifyAttribute,T>();

            var excelSheet = dataExport.CreateSheet(extractedAttributes);

            using var fileStream = new FileStream($"{fileName}.xlsx", FileMode.Create, FileAccess.Write);
           
            excelSheet.Write(fileStream);

            return fileStream;
        }

        public override XSSFWorkbook ExportToExcel<T>(IEntityExport<T> dataExport)
        {
            var extractedAttributes = ExcelifyRecord.GetAttributeProperty<ExcelifyAttribute, T>();

            var excelSheet = dataExport.CreateSheet(extractedAttributes);

            return excelSheet;
        }

        private int _sheetName;
        private readonly ExcelifyMapper _excelifyMapper = new();
    }
}
