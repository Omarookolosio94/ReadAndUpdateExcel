using System.Data;

namespace ExcelReader.Services.Interfaces
{
    public interface IProductRepository
    {
        bool SaveUpdatedProductRecordsToDatabase(DataTable excelData);
    }
}
