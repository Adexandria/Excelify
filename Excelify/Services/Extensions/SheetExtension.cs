using Excelify.Models;
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
    }
}
