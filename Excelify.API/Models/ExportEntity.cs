using Excelify.Models;

namespace Excelify.API.Models
{
    public class ExportEntity : IEntityExport<TeacherDTO>
    {
        public string SheetName { get; set ; }
        public IList<TeacherDTO> Entities { get; set ; }
    }
}
