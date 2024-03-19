using Excelify.Models;

namespace Excelify.API.Models
{
    public class ImportSheet : IImportSheet
    {
        public ImportSheet(Stream file)
        {
            File = file;
        }
        public Stream File { get; set; }
    }
}
