namespace ExcelReader.Services.Interfaces
{
    public interface IEmailService
    {
        void SendEmail(Message message);
        void SendHTMLEmail(Message message);
        void SendEmailWithAttachment(Message message, string fileName);
        Task SendEmailAsync(Message message);
    }
}
