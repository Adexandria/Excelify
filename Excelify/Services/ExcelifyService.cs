﻿using Excelify.Models;
using Excelify.Services.Extensions;
using Excelify.Services.Utility;
using Excelify.Services.Utility.Attributes;
using System.ComponentModel;
using System.Data;

namespace Excelify.Services
{
    public class ExcelifyService : ExcelService
    {

        public override void SetSheetName(int sheetName , string extensionType)
        {
            if(string.IsNullOrEmpty(extensionType))
                throw new ArgumentNullException(nameof(extensionType),"Extension type can not be empty");

            if (sheetName < 0)
                throw new ArgumentNullException(nameof(sheetName), "Invalid sheet name");


            if(extensionType.Equals(ExtensionType.xls.GetDescription<DescriptionAttribute>()) ||
                extensionType.Equals(ExtensionType.xlsx.GetDescription<DescriptionAttribute>()))
            {
                _sheetName = sheetName;
            }
            else
            {
                throw new Exception("Invalid extension type", new Exception("Only xls and xlsx are accepted"));
            }
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

        public override byte[] ExportToExcel<T>(IEntityExport<T> dataExport)
        {
            var extractedAttributes = ExcelifyRecord.GetAttributeProperty<ExcelifyAttribute, T>();

            var excelSheet = dataExport.CreateSheet(extractedAttributes);

            using var memoryStream = new MemoryStream();

            excelSheet.Write(memoryStream);

            return memoryStream.ToArray();
        }

        private int _sheetName;
        private readonly ExcelifyMapper _excelifyMapper = new();
    }
}
