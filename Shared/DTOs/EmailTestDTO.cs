
namespace IMPLANPROD.Shared.DTOs
{
    public class EmailTestDTO
    {
        public EmailConfigDTO Config { get; set; }
        public byte[] AttachmentContent { get; set; }
        public string AttachmentName { get; set; }
        public string AttachmentContentType { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
    }
}