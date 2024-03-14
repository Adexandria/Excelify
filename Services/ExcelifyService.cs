﻿using Excelify.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        public override Task<DataTable> ImportSheet(IExcelifySheet sheet)
        {
            throw new NotImplementedException();
        }

        public override Task<IList<T>> ImportToEntity<T>(IExcelifySheet sheet)
        {
            throw new NotImplementedException();
        }

        private DataTable ExtractValuesFromSheet(IExcelifySheet excelSheet)
        {
            DataTable table = new();
            List<string> rowList = new();
            ISheet sheet;
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

        private readonly int _sheetName;
    }
}
