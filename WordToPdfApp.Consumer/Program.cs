using System;
using Spire.Doc;
using System.IO;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System.Net.Mail;
using System.Net.Mime;
using RabbitMQ.Client.Events;

namespace WordToPdfApp.Consumer
{
    internal class Program
    {
        private static void Main()
        {
            bool result = false;
            var factory = new ConnectionFactory();
            factory.Uri = new Uri("amqps://zbwwvfdt:Zx52kZt2NEYXU0RvsoJiuW7Ar5W-iA6V@woodpecker.rmq.cloudamqp.com/zbwwvfdt");
            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);
                    channel.QueueBind("File", "convert-exchange", "WordToPdf");
                    channel.BasicQos(0, 1, false);

                    var consumer = new EventingBasicConsumer(channel);
                    channel.BasicConsume("File", false, consumer);
                    consumer.Received += (model, ea) =>
                    {
                        try
                        {
                            Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor");
                            Document document = new();
                            string message = Encoding.UTF8.GetString(ea.Body.ToArray());
                            MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(message);
                            document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Docx2013);
                            using (MemoryStream memoryStream = new())
                            {
                                document.SaveToStream(memoryStream, FileFormat.PDF);
                                result = EmailSend(messageWordToPdf.Email, memoryStream, messageWordToPdf.FileName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Hata meydana geldi: " + ex.Message);
                        }
                        if (result)
                        {
                            Console.WriteLine("Kuyruktan mesaj başarıyla işlendi...");
                            channel.BasicAck(ea.DeliveryTag, false);
                        }
                    };
                    Console.WriteLine("Çıkmak için tıklayınız");
                    Console.ReadLine();
                }
            }
        }

        public static bool EmailSend(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0;
                ContentType contentType = new(MediaTypeNames.Application.Pdf);
                Attachment attachment = new(memoryStream, contentType);
                attachment.ContentDisposition.FileName = $"{fileName}.pdf";
                MailMessage mailMessage = new();
                SmtpClient smtpClient = new();
                mailMessage.From = new MailAddress("admin@teknohub.net");
                mailMessage.To.Add(email);
                mailMessage.Subject = "Pdf Dosyası Oluşturma | teknohub.net";
                mailMessage.Body = "Pdf dosyanız ektedir.";
                mailMessage.IsBodyHtml = true;
                mailMessage.Attachments.Add(attachment);

                smtpClient.Host = "mail.teknohub.net";
                smtpClient.Port = 587;
                smtpClient.Credentials = new NetworkCredential("admin@teknohub.net", "Fatih1234");
                smtpClient.Send(mailMessage);
                Console.WriteLine($"Sonuç {email} adresine gönderilmiştir.");
                memoryStream.Close();
                memoryStream.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Mail gönderim sırasında bir hata meydana geldi: {ex.InnerException}");
                return false;
            }
        }
    }
}