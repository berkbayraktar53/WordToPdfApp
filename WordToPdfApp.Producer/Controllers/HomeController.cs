using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using RabbitMQ.Client;
using Microsoft.AspNetCore.Mvc;
using WordToPdfApp.Producer.Models;
using Microsoft.Extensions.Configuration;

namespace WordToPdfApp.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;

        public HomeController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult WordToPdfPage()
        {
            return View();
        }

        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            var factory = new ConnectionFactory
            {
                Uri = new Uri(_configuration["ConnectionStrings:RabbitMQCloudString"])
            };
            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);
                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, arguments: null);
                    channel.QueueBind("File", "convert-exchange", "WordToPdf");

                    MessageWordToPdf messageWordToPdf = new();
                    using (MemoryStream memoryStream = new())
                    {
                        wordToPdf.WordFile.CopyTo(memoryStream);
                        messageWordToPdf.WordByte = memoryStream.ToArray();
                    }
                    messageWordToPdf.Email = wordToPdf.Email;
                    messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                    string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);
                    byte[] byteMessage = Encoding.UTF8.GetBytes(serializeMessage);
                    var properties = channel.CreateBasicProperties();
                    properties.Persistent = true;
                    channel.BasicPublish("convert-exchange", routingKey: "WordToPdf", basicProperties: properties, body: byteMessage);

                    ViewBag.result = "Word dosyanız pdf dosyasına dönüştürüldükten sonra size email olarak gönderilecektir.";

                    return View();
                }
            }
        }
    }
}