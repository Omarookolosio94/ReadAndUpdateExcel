using MimeKit;

namespace ExcelReader.Models
{
    public class Message
    {
        public List<MailboxAddress> To { get; set; }
        public string Subject { get; set; }
        public string Content { get; set; }
        public byte[] Attachments { get; set; }
        public Message(IEnumerable<string> to, string subject, string content , byte[] pdfFile = null)
        {
            To = new List<MailboxAddress>();
            To.AddRange(to.Select(y => MailboxAddress.Parse(y)));
            Subject = subject;
            Content = content;
            Attachments = pdfFile;
        }
    }
}
