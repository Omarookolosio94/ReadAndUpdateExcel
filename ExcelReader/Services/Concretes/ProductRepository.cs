using System.Data;
using System.Data.SqlClient;

namespace ExcelReader.Services.Concretes
{
    public class ProductRepository : IProductRepository
    {
        public bool SaveUpdatedProductRecordsToDatabase(DataTable excelData)
        {
            try
            {
                using var connection = new SqlConnection(Appsettings.ConnectionString);

                connection.Open();

                using (var bulk = new SqlBulkCopy(connection.ConnectionString, SqlBulkCopyOptions.Default | SqlBulkCopyOptions.TableLock))
                {
                    bulk.DestinationTableName = "Product";

                    SqlBulkCopyColumnMapping ProductId = new("ProductId", "ProductId");
                    bulk.ColumnMappings.Add(ProductId);

                    SqlBulkCopyColumnMapping Store = new("Store", "Store");
                    bulk.ColumnMappings.Add(Store);

                    SqlBulkCopyColumnMapping UnitPrice = new("UnitPrice", "UnitPrice");
                    bulk.ColumnMappings.Add(UnitPrice);

                    SqlBulkCopyColumnMapping Name = new("Name", "Name");
                    bulk.ColumnMappings.Add(Name);

                    SqlBulkCopyColumnMapping Summary = new("Summary", "Summary");
                    bulk.ColumnMappings.Add(Summary);

                    SqlBulkCopyColumnMapping Category = new("Category", "Category");
                    bulk.ColumnMappings.Add(Category);

                    SqlBulkCopyColumnMapping Brand = new("Brand", "Brand");
                    bulk.ColumnMappings.Add(Brand);

                    SqlBulkCopyColumnMapping IsProductAvailable = new("IsProductAvailable", "IsProductAvailable");
                    bulk.ColumnMappings.Add(IsProductAvailable);

                    bulk.BulkCopyTimeout = 0;
                    bulk.BatchSize = 45;
                    bulk.EnableStreaming = true;
                    bulk.WriteToServer(excelData);
                }

                connection.Close();

                return true;

            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}