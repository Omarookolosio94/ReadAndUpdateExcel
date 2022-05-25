using ExcelReader.DAO;
using ExcelReader.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Web;

namespace ExcelReader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReadExelController : ControllerBase
    {
        [HttpPost("readAndUpdateExcel")]
        public async Task<IActionResult> ReadAndUpdateExcel(IFormFile formFile, CancellationToken cancellationToken)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //validate file
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest(new BaseResponse<string>(400, "Please upload an excel file", ""));
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new BaseResponse<string>(400, "File format not supported, Please upload an excel file", ""));
            }

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

            using var stream = new MemoryStream();

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
                            return BadRequest(new BaseResponse<string>(400, "An error occured when processing request", ex.ToString()));

                        }
                    }

                    var downloadStream = new MemoryStream();

                    using (ExcelPackage downloadPackage = new(downloadStream))
                    {
                        ExcelWorksheet downloadWorksheet = downloadPackage.Workbook.Worksheets.Add("UpdatedBasketApi");
                        downloadWorksheet.Cells["A1"].LoadFromDataTable(updatedExcel, true);
                        downloadPackage.Save();
                    }

                    downloadStream.Position = 0;

                    string excelName = $"UpdatedBasketApiList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

                    return File(downloadStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
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
