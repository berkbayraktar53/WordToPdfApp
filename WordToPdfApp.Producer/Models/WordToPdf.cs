using Microsoft.AspNetCore.Http;

namespace WordToPdfApp.Producer.Models
{
    public class WordToPdf
    {
        public string Email { get; set; }
        public IFormFile WordFile { get; set; }
    }
}