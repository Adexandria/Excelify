using Excelify.Models;

namespace Excelify.API.Models
{
    public class ImportSheet(Stream file) : IImportSheet
    {
        public Stream File { get; set; } = file;
    }
}
