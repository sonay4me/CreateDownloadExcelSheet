using CreateExcelSheet.Data;
using CreateExcelSheet.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace CreateExcelSheet.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelGenerationService _excelService;
        private readonly IStudentService _studentService;
        public ExcelController(IExcelGenerationService excelService, IStudentService studentService)
        {
            _excelService = excelService;
            _studentService = studentService;
        }

        [HttpGet]
        public async Task<ActionResult> GetExcel()
        {
            var students = _studentService.Students();
            MemoryStream memoryStream = await _excelService.GenerateStudentList(students);
            FileStreamResult fileStreamResult = new(memoryStream, "application/xls")
            {
                FileDownloadName = "StudentList.xlsx"
            };
            return fileStreamResult;
        }
    }
}
