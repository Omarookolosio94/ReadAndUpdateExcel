
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;
using static ExcelReader.DAO.FetchDataDAO;

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

        public async Task UpdateRecordProcess(string? streamJson, string email)
        {
            try
            {
                //convert json string back to Memory stream
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

                    string backupTableName = "";

                    //create a new backup of table
                    bool backupTableStatus = BackUpTableandTruncate("Product", false, out backupTableName);

                    if (backupTableStatus)
                    {
                        try
                        {
                            //save to database
                            var response = _productRepository.SaveUpdatedProductRecordsToDatabase(updatedExcel);

                            if (response)
                            {
                                var message = new Message(new string[] { email }, "Updating Data Successfull", "<h1>Updating Excel Records Updated Successfully</h1><p>Data was updated in both excel and database</p>", downloadStream.ToArray());
                                _emailService.SendEmailWithAttachment(message, excelName);
                            }
                            else
                            {
                                bool revertStatus = BackUpTableandTruncate(backupTableName, true, out backupTableName);

                                if (revertStatus)
                                {
                                    var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Updated But Database Update Failed</h1><p>Data was updated in excel file but could not be uploaded to the database. Table was reverted back successfully</p>", downloadStream.ToArray());
                                    _emailService.SendEmailWithAttachment(message, excelName);
                                }
                                else
                                {
                                    var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Updated But Database Update Failed</h1><p>Data was updated in excel file but could not be uploaded to the database. Table reversion failed, Kindly engage the support team to manually revert the table</p>", downloadStream.ToArray());
                                    _emailService.SendEmailWithAttachment(message, excelName);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Updated But Database Update Failed</h1><p>Data was updated in excel file but could not be uploaded to the database. Table backup couldnot be created</p>", downloadStream.ToArray());
                            _emailService.SendEmailWithAttachment(message, excelName);
                        }
                    }
                    else
                    {
                        var message = new Message(new string[] { email }, "Updating Data Failure", "<h1>Updating Excel Records Updated But Database Update Failed</h1><p>Data was updated in excel file but could not be uploaded to the database. Table backup couldnot be created</p>", downloadStream.ToArray());
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

        public bool BackUpTableandTruncate(string existingTable, bool revert, out string backupTableName)
        {
            backupTableName = existingTable + DateTime.Now.ToString("ddMMMyyyyHHmm");

            string existingTableToBackup = "[excelreader].[dbo].[" + existingTable + "]";

            string queryString;

            if (!revert)
            {
                queryString = string.Format("Select * into [" + backupTableName + "] from " + existingTableToBackup + "GO Delete From " + existingTableToBackup);
            }
            else
            {
                backupTableName = "Product";
                queryString = string.Format("Delete From " + backupTableName + ";Insert into [" + backupTableName + "] select ProductId,Store,UnitPrice,Name,Summary,Category,Brand,IsProductAvailable from " + existingTableToBackup);
            }

            try
            {
                using (SqlConnection sqlConnection = new(Appsettings.ConnectionString))
                {
                    SqlCommand sqlCommand = new(string.Format(queryString.Trim()), sqlConnection);

                    sqlConnection.Open();

                    using SqlDataReader reader = sqlCommand.ExecuteReader();

                    while (reader.Read())
                    {
                        Console.WriteLine(String.Format("{0}, {1}",
                            reader[0], reader[1]));
                    }
                }

                return true;

            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
}
