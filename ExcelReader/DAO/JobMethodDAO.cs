
using System.Data;

namespace ExcelReader.DAO
{
    public class JobMethodDAO
    {
        private readonly IProductRepository _productRepository;
        private readonly IEmailService _emailService;
        public JobMethodDAO(IProductRepository productRepository, IEmailService emailService)
        {
            _productRepository = productRepository;
            _emailService = emailService;
        }
        public void UpdateRecordProcess(ExcelWorksheet excelworksheet, string email)
        {
            var worksheet = excelworksheet;

            try
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

                int lastRow = worksheet.Dimension.End.Row;

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
            catch (Exception ex)
            {
                //Email to user that process failed
                var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Failed</h1><p>Data couldnot be extracted from the excel document</p>");
                _emailService.SendHTMLEmail(message);
            }
        }

    }
}
