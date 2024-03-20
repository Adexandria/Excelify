using Excelify.Models;
using Excelify.Services.Utility;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace Excelify.Services.Extensions
{
    internal static class SheetExtension
    {
        public static DataTable ExtractValues(this IImportSheet excelSheet, int _sheetName)
        {
            ISheet sheet;
            DataTable table = new();
            List<string> rowList = [];
            XSSFWorkbook workBook = new(excelSheet.File)
            {
                MissingCellPolicy = MissingCellPolicy.RETURN_NULL_AND_BLANK
            };

            sheet = workBook.GetSheetAt(_sheetName);

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            for (int i = 0; i < cellCount; i++)
            {
                ICell cell = headerRow.GetCell(i);
                if (cell == null || string.IsNullOrEmpty(cell.ToString()))
                    continue;

                table.Columns.Add(cell.ToString());
            }
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank))
                    continue;
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    var cell = row.GetCell(j);

                    if (cell == null)
                    {
                        rowList.Add(null);
                    }
                    else
                    {
                        rowList.Add(cell.ToString());
                    }
                }

                if (rowList.Count > 0)
                    table.Rows.Add(rowList.ToArray());

                rowList.Clear();
            }
            return table;
        }

        public static XSSFWorkbook CreateSheet<T>(this IEntityExport<T> dataExport, List<ExcelifyProperty> extractedAttributes)
        {
            var workBook = new XSSFWorkbook();
            var workSheet = workBook.CreateSheet(dataExport.SheetName);
            var headerRow = workSheet.CreateRow(0);
            for (int i = 0; i < extractedAttributes.Count; i++)
            {
                ICell cell;
                if (extractedAttributes[i].AttributeName is int value)
                {
                    headerRow.CreateCell(value, CellType.String)
                       .SetCellValue(extractedAttributes[i].PropertyName);
                    InsertValues(workSheet, dataExport.Entities,
                        extractedAttributes[i].PropertyName, i);
                }
                else
                {
                    headerRow.CreateCell(i, CellType.String)
                         .SetCellValue((string)extractedAttributes[i].AttributeName);
                    InsertValues(workSheet, dataExport.Entities,
                       extractedAttributes[i].PropertyName, i);
                }
            }

            return workBook;
        }

        public static void WriteToFile(this byte[] workbook, string fileName)
        {
            if(string.IsNullOrEmpty(fileName) || string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException(nameof(fileName), "File name can not be empty");
            }

            using var fileStream = new FileStream($"{fileName}.xlsx", FileMode.Create, FileAccess.Write);

            fileStream.Write(workbook,0, workbook.Length);
        }

        private static void InsertValues<T>(ISheet sheet, IList<T> entities, string propertyName, int cellNumber, int rowNumber = 1)
        {
            if (rowNumber <= entities.Count)
            {
                var column = sheet.GetRow(rowNumber) ?? sheet.CreateRow(rowNumber);
                var entity = entities[rowNumber - 1];
                var property = entity.GetType().GetProperty(propertyName);
                if (property == null)
                {
                    return;
                }

                switch (property.PropertyType )
                {
                    case Type when property.PropertyType == typeof(int) :
                        column.CreateCell(cellNumber, CellType.Numeric)
                         .SetCellValue((int)property.GetValue(entity));
                    break;

                    case Type when property.PropertyType == typeof(double):
                        column.CreateCell(cellNumber, CellType.Numeric)
                         .SetCellValue((double)property.GetValue(entity));
                    break;

                    case Type when property.PropertyType == typeof(DateTime):
                        var value = (DateTime)property.GetValue(entity);

                        column.CreateCell(cellNumber, CellType.String)
                            .SetCellValue(value.ToString());
                        break;

                    case Type when property.PropertyType == typeof(bool):
                        column.CreateCell(cellNumber, CellType.Boolean)
                             .SetCellValue((bool)property.GetValue(entity));
                        break;
                        
                    default:
                        var x = column.CreateCell(cellNumber, CellType.String);
                        x.SetCellValue((string)property.GetValue(entity));
                        x.CellStyle.WrapText = true;
                        break;
                }

                InsertValues(sheet, entities, propertyName, cellNumber, ++rowNumber);
            }
        }
    }
}
