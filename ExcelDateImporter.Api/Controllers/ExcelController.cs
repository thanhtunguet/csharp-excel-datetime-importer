using Microsoft.AspNetCore.Mvc;
using ExcelDateImporter.Api.Services;
using ExcelDateImporter.Api.Models;

namespace ExcelDateImporter.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ExcelDateParsingService _excelService;

        public ExcelController(ExcelDateParsingService excelService)
        {
            _excelService = excelService;
        }

        [HttpPost("upload")]
        public async Task<ActionResult<List<ExcelDateEntry>>> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded");
            }

            var allowedExtensions = new[] { ".xlsx", ".xls" };
            var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
            
            if (!allowedExtensions.Contains(fileExtension))
            {
                return BadRequest("Only Excel files (.xlsx, .xls) are allowed");
            }

            try
            {
                var results = await _excelService.ProcessExcelFile(file);
                return Ok(results);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error processing file: {ex.Message}");
            }
        }

        [HttpGet("code-example")]
        public ActionResult<string> GetCodeExample()
        {
            var codeExample = _excelService.GetCSharpCodeExample();
            return Ok(codeExample);
        }
    }
}