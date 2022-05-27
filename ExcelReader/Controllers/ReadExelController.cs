using ExcelReader.DAO;
using Hangfire;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Data;
using static ExcelReader.DAO.FetchDataDAO;

namespace ExcelReader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReadExelController : ControllerBase
    {
        private readonly IBackgroundJobClient _backgroundJobClient;
        private readonly IProductRepository _productRepository;
        private readonly IEmailService _emailService;
        public ReadExelController(IProductRepository productRepository, IEmailService emailService, IBackgroundJobClient backgroundJobClient)
        {
            _backgroundJobClient = backgroundJobClient;
            _productRepository = productRepository;
            _emailService = emailService;
        }

        [HttpPost("readAndUpdateExcel")]
        public async Task<IActionResult> ReadAndUpdateExcel(IFormFile formFile, [FromQuery] string email, CancellationToken cancellationToken)
        {
            //License attribute for EPPlus extension used in reading excel file
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //validation of input parameters
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest(new BaseResponse<string>(400, "Please upload an excel file", ""));
            }

            if (string.IsNullOrEmpty(email))
            {
                return BadRequest(new BaseResponse<string>(400, "Please include a valid email address", ""));
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new BaseResponse<string>(400, "File format not supported, Please upload an excel file", ""));
            }

            using var stream = new MemoryStream();

            string excelName = $"UpdatedBasketApiList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

            await formFile.CopyToAsync(stream, cancellationToken);

            using var package = new ExcelPackage(stream);

            ExcelWorksheet worksheet = package.Workbook.Worksheets["BasketApi"];

            if (worksheet != null)
            {
                int lastRow = worksheet.Dimension.End.Row;
                while (worksheet.Cells[lastRow, 1].Value == null)
                {
                    lastRow--;
                }

                //Check if all column headers are named properly
                if (worksheet.Cells[1, 1].Value.Equals("ProductId") &&
                    worksheet.Cells[1, 2].Value.Equals("Store"))
                {
                    //convert stream to jsonstring reason been the background job cannot read stream data and IFormfile directly 
                    var streamJson = JsonConvert.SerializeObject(stream, Formatting.Indented, new MemoryStreamJsonConverter());

                    _backgroundJobClient.Enqueue(() => new JobMethodDAO(_productRepository , _emailService).UpdateRecordProcess(streamJson, email));

                    return Ok(new BaseResponse<string>(200, "Records from excel is been updated at the moment. You will get an email on the progress and status when it is completed", ""));
                }
                else
                {
                    return BadRequest(new BaseResponse<string>(400, "Please upload excel with the specified format: ProductId, Store", ""));
                }
            }
            else
            {
                return BadRequest(new BaseResponse<string>(400, "Please add the workspace 'BasketApi' in the upload excel file", ""));
            }

        }
    }
}
