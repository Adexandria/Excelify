using Excelify.Models;
using Excelify.Services.Extensions;
using Excelify.Services.Utility;
using Excelify.Services.Utility.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace Excelify.Services
{
    public class ExcelifyService(int sheetName = 0) : ExcelService
    {
        public override DataTable ImportSheet(IImportSheet sheet)
        {
           return sheet.ExtractValues(_sheetName);
        }

        public override IList<T> ImportToEntity<T>(IImportSheet sheet)
        {
            var extractedValues = sheet.ExtractValues(_sheetName);
            var entities = _excelifyMapper.Map<T>(extractedValues.Rows.OfType<DataRow>()).Result;
            return entities;
        }

        public override IList<T> ImportToEntity<T>(IImportSheet sheet, IExcelMapper excelifyMapper)
        {
            var extractedValues = sheet.ExtractValues(_sheetName);
            var entities = excelifyMapper.Map<T>(extractedValues.Rows.OfType<DataRow>()).Result;
            return entities;
        }

        public override ISheet ExportToExcel<T>(IEntityExport<T> dataExport)
        {
            var extractedAttributes = ExcelifyRecord.GetAttributeProperty<ExcelifyAttribute,T>();
            var excelSheet = CreateSheet(dataExport, extractedAttributes);
            return excelSheet;
        }

        private ISheet CreateSheet<T>(IEntityExport<T> dataExport,List<ExcelifyProperty> extractedAttributes)
        {
            var workSheet = new XSSFWorkbook().CreateSheet(dataExport.SheetName);
            var headerRow =  workSheet.CreateRow(0);
            for(int i = 0; i < extractedAttributes.Count; i++)
            {
                ICell cell;
                if (extractedAttributes[i].AttributeName is int value)
                {
                     headerRow.CreateCell(value, CellType.String)
                        .SetCellValue(extractedAttributes[i].PropertyName);
                    InsertValues(workSheet, dataExport.Entities, 
                        extractedAttributes[i].PropertyName, extractedAttributes[i].PropertyType, i);
                }
                else
                {
                   headerRow.CreateCell(i, CellType.String)
                        .SetCellValue((string)extractedAttributes[i].AttributeName);
                    InsertValues(workSheet, dataExport.Entities,
                       extractedAttributes[i].PropertyName, extractedAttributes[i].PropertyType, i);
                }
            }

            return workSheet;
        }

        private void InsertValues<T>(ISheet sheet, IList<T> entities,string propertyName,object propertyType, int cellNumber, int numberOfRows = 1)
        {
            var column = sheet.CreateRow(numberOfRows);
            if(numberOfRows <= entities.Count)
            {
               var entity = entities[numberOfRows];
               var property = entity.GetType().GetProperty(propertyName);
               switch (propertyType)
                {
                    case string:
                        column.CreateCell(cellNumber, CellType.String)
                            .SetCellValue((string)property.GetValue(entity));
                        break;
                    case int:
                        column.CreateCell(cellNumber, CellType.Numeric)
                            .SetCellValue((int)property.GetValue(entity));
                        break;
                    case double:
                        column.CreateCell(cellNumber, CellType.Numeric)
                            .SetCellValue((double)property.GetValue(entity));
                        break;
                    case float:
                        column.CreateCell(cellNumber, CellType.Numeric)
                             .SetCellValue((double)property.GetValue(entity));
                        break;
                    case DateTime:
                        column.CreateCell(cellNumber, CellType.Unknown)
                            .SetCellValue((DateTime)property.GetValue(entity));
                        break;
                    case bool:
                        column.CreateCell(cellNumber, CellType.Boolean)
                             .SetCellValue((bool)property.GetValue(entity));
                        break;
                    default:
                        column.CreateCell(cellNumber, CellType.Unknown)
                             .SetCellValue((string)property.GetValue(entity));
                        break;
               }
                InsertValues(sheet, entities, propertyName, propertyType, cellNumber, numberOfRows++);
            }
        }

        private readonly int _sheetName = sheetName;
        private readonly ExcelifyMapper _excelifyMapper = new();
    }
}
