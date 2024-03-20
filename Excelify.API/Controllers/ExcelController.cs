using Excelify.API.Models;
using Excelify.Services;
using Excelify.Services.Extensions;
using Microsoft.AspNetCore.Mvc;

namespace Excelify.API.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        public ExcelController(IExcelService excelService)
        {
            _excelService = excelService;      
        }
        [HttpPost("import/{sheetName}")]
        public  IActionResult ImportExcel(IFormFile sheet,int sheetName = 0)
        {
            _excelService.SetSheetName(sheetName,sheet.ContentType);
            try
            {
                var ms = new MemoryStream();
                sheet.CopyTo(ms);
                ms.Position = 0;

                var teacherDtos = _excelService.ImportToEntity<TeacherDTO>(new ImportSheet(ms));

                if (teacherDtos.Count == 0)
                {
                    return BadRequest("Empty excel sheet");
                }

                var teachers = teacherDtos.Select(s => new Teacher
                {
                    StartDate = s.StartDate,
                    PhoneNumber = s.PhoneNumber,
                    Email = s.Email,
                    FirstName = s.FirstName,
                    Id = s.Id,
                    LastName = s.LastName,
                    Title = s.Title
                });
                return Ok(teachers);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("export/{sheetName}")]
        public IActionResult ExportToExcel(string sheetName)
        {
            
            var teacher = new TeacherDTO() 
            {
                Email = "adeolaaderibigbe09@gmail.com",
                PhoneNumber = 0000,
                FirstName = "Ade",
                LastName = "ad",
                Id = 1,
                StartDate = DateTime.Now,
                Title = "Miss"
            };
            var exportEntity = new ExportEntity()
            { 
                Entities = new[] { teacher },
                SheetName = sheetName
            };

            try
            {
                _excelService.ExportToExcel(exportEntity)
                    .ToFile($"{sheetName}.xlsx");
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private readonly IExcelService _excelService;
    }
}
