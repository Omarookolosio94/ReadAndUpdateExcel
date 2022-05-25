namespace ExcelReader.Models
{
    public class Product
    {
        public int ProductID { get; set; }
        public int UnitPrice { get; set; }
        public string? Name { get; set; }
        public string? Summary { get; set; }
        public string? Category { get; set; }
        public string? Brand { get; set; }
    }

    public class BaseResponse<T>
    {
        public BaseResponse(int statusCode, string message, T data)
        {
            Code = statusCode;
            Message = message;
            Data = data;
        }
        public int Code { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }
        
    }
}
