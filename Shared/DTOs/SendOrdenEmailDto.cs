namespace IMPLANPROD.Shared.DTOs
{
    public class SendOrdenEmailDto
    {
        public string To { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public byte[]? AttachmentContent { get; set; }
        public string? AttachmentName { get; set; }
        public string? ReplyTo { get; set; }
        public string? ReplyToName { get; set; }
    }
}
