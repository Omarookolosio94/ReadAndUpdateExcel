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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //validate file
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
            var uploadedFile = File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            
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
                    var streamJson = JsonConvert.SerializeObject(uploadedFile, Formatting.Indented, new MemoryStreamJsonConverter());

                    _backgroundJobClient.Enqueue(() => UpdateRecordProcess(streamJson, email));

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

        [HttpGet]
        public async Task UpdateRecordProcess(string? streamJson, string email)
        {
            try
            {
                var stream = JsonConvert.DeserializeObject<MemoryStream>(streamJson, new MemoryStreamJsonConverter());

                using var package = new ExcelPackage(stream);

                ExcelWorksheet worksheet = package.Workbook.Worksheets["BasketApi"];

                int lastRow = worksheet.Dimension.End.Row;

                while (worksheet.Cells[lastRow, 1].Value == null)
                {
                    lastRow--;
                }

                //Check if all column headers are named properly
                if (worksheet.Cells[1, 1].Value.Equals("ProductId") &&
                    worksheet.Cells[1, 2].Value.Equals("Store"))
                {
                    //extract data from excel
                    DataTable updatedExcel = new();

                    updatedExcel.Columns.Add(new DataColumn("ProductId"));
                    updatedExcel.Columns.Add(new DataColumn("Store"));
                    updatedExcel.Columns.Add(new DataColumn("UnitPrice"));
                    updatedExcel.Columns.Add(new DataColumn("Name"));
                    updatedExcel.Columns.Add(new DataColumn("Summary"));
                    updatedExcel.Columns.Add(new DataColumn("Category"));
                    updatedExcel.Columns.Add(new DataColumn("Brand"));
                    updatedExcel.Columns.Add(new DataColumn("IsProductAvailable"));

                    for (int startRow = 2; startRow <= lastRow; startRow++)
                    {
                        try
                        {
                            var productItemId = Convert.ToInt32(worksheet.Cells[startRow, 1].Value.ToString());
                            var store = worksheet.Cells[startRow, 2].Value.ToString();

                            //fetch record from API
                            var fetchedProduct = new FetchDataDAO().GetProduct(productItemId);
                            var fetchRatings = new FetchDataDAO().GetRatings(productItemId);

                            DataRow row = updatedExcel.NewRow();

                            if (fetchedProduct != null && fetchRatings != null)
                            {

                                row["ProductId"] = productItemId;
                                row["Store"] = store;
                                row["UnitPrice"] = fetchedProduct.UnitPrice;
                                row["Name"] = fetchedProduct.Name;
                                row["Summary"] = fetchedProduct.Summary;
                                row["Category"] = fetchedProduct.Category;
                                row["Brand"] = fetchedProduct.Brand;
                                row["IsProductAvailable"] = fetchRatings.IsAvailable;

                                updatedExcel.Rows.Add(row);

                            }
                            else if (fetchedProduct != null && fetchRatings == null)
                            {
                                row["ProductId"] = productItemId;
                                row["Store"] = store;
                                row["UnitPrice"] = fetchedProduct.UnitPrice;
                                row["Name"] = fetchedProduct.Name;
                                row["Summary"] = fetchedProduct.Summary;
                                row["Category"] = fetchedProduct.Category;
                                row["Brand"] = fetchedProduct.Brand;
                                row["IsProductAvailable"] = false;

                                updatedExcel.Rows.Add(row);

                            }
                            else if (fetchRatings != null && fetchedProduct == null)
                            {
                                row["ProductId"] = productItemId;
                                row["Store"] = store;
                                row["UnitPrice"] = 0;
                                row["Name"] = "";
                                row["Summary"] = "";
                                row["Category"] = "";
                                row["Brand"] = "";
                                row["IsProductAvailable"] = fetchRatings.IsAvailable;

                                updatedExcel.Rows.Add(row);
                            }
                            else if (fetchedProduct == null && fetchedProduct == null)
                            {
                                row["ProductId"] = productItemId;
                                row["Store"] = store;
                                row["UnitPrice"] = 0;
                                row["Name"] = "";
                                row["Summary"] = "";
                                row["Category"] = "";
                                row["Brand"] = "";
                                row["IsProductAvailable"] = false;

                                updatedExcel.Rows.Add(row);
                            }

                        }
                        catch (Exception ex)
                        {
                            var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Failed</h1><p>Data couldnot be extracted from the excel document</p>");
                            _emailService.SendHTMLEmail(message);
                        }
                    }

                    //send file to user email
                    var downloadStream = new MemoryStream();
                    using (ExcelPackage downloadPackage = new(downloadStream))
                    {
                        ExcelWorksheet downloadWorksheet = downloadPackage.Workbook.Worksheets.Add("UpdatedBasketApi");
                        downloadWorksheet.Cells["A1"].LoadFromDataTable(updatedExcel, true);
                        downloadPackage.Save();
                    }

                    downloadStream.Position = 0;

                    string excelName = $"UpdatedBasketApiList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
                    //var excelFile = File(downloadStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);

                    //save to database
                    var response = _productRepository.SaveUpdatedProductRecordsToDatabase(updatedExcel);

                    if (response)
                    {
                        var message = new Message(new string[] { email }, "Updating Data Successfull", "<h1>Updating Excel Records Updated Successfully</h1><p>Data was updated in bothh excel and database</p>", downloadStream.ToArray());
                        _emailService.SendEmailWithAttachment(message, excelName);
                    }
                    else
                    {
                        var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Updated But Database Update Failed</h1><p>Data was updated in excel file but could not be uploaded to the database</p>", downloadStream.ToArray());
                        _emailService.SendEmailWithAttachment(message, excelName);
                    }
                }
                else
                {
                    var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Failed</h1><p>Data couldnot be extracted from the excel document</p>");
                    _emailService.SendHTMLEmail(message);
                }
            }
            catch (Exception ex)
            {
                //Email to user that process failed
                var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Failed</h1><p>Data couldnot be extracted from the excel document</p>");
                _emailService.SendHTMLEmail(message);
            }
        }

    }
}
