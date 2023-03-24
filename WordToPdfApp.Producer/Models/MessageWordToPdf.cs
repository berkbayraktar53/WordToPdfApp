namespace WordToPdfApp.Producer.Models
{
    public class MessageWordToPdf
    {
        public string Email { get; set; }
        public string FileName { get; set; }
        public byte[] WordByte { get; set; }
    }
}