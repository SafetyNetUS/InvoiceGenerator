using CsvHelper;
using CsvHelper.Configuration.Attributes;
using MailKit.Net.Smtp;
using MimeKit;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace InvoiceGenerator
{
    public class User
    {
        [Name("BTC UBI Subscribed")]
        public string BtcSubscribed { get; set; }
        [Name("BTC Wallet")]
        public string BtcWallet { get; set; }
        [Name("Communication Email")]
        public string CommunicationEmail { get; set; }
        [Name("USD UBI Subscribed")]
        public string UsdSubscribed { get; set; }
        [Name("Creation Date")]
        public DateTime CreationDate { get; set; }
        [Name("email")]
        public string Email { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            const int FirstInvoiceNumberOfBatch = 29;
            const string UsersPath = @"/path/to/list/of/users";
            const string InvoiceDirectory = @"/path/to/generated/invoices";
            const string SmtpHost = "The SMTP Server host";
            const int SmtpPort = 465;
            const string SmtpUsername = "The SMTP user name";
            const string SmtpPassword = "The SMTP password";

            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("invoicegenerator.log")
                .WriteTo.Console()
                .CreateLogger();

            using var reader = new StreamReader(UsersPath);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            var currentInvoiceNumber = FirstInvoiceNumberOfBatch;

            using var smtpClient = new SmtpClient();
            await smtpClient.ConnectAsync(SmtpHost, SmtpPort, true);
            await smtpClient.AuthenticateAsync(SmtpUsername, SmtpPassword);

            await foreach (var user in csv.GetRecordsAsync<User>())
            {
                var invoiceFilePath = await GenerateInvoiceAsync(user.CommunicationEmail, user.Email, currentInvoiceNumber++, InvoiceDirectory);
                await SendEmailAsync(user.CommunicationEmail, invoiceFilePath, smtpClient);
                Log.Information($"Processed invoice #{currentInvoiceNumber - 1} - {user.Email}");
            }

            await smtpClient.DisconnectAsync(true);
        }

        private static async Task SendEmailAsync(string recipientEmail, string attachmentFilePath, SmtpClient smtpClient)
        {
            var mailMessage = new MimeMessage();
            mailMessage.From.Add(new MailboxAddress("Safety Net", "help@safetynetus.io"));
            mailMessage.To.Add(new MailboxAddress(recipientEmail, recipientEmail));
            mailMessage.Subject = "December 2020 UBI Distribution";

            var builder = new BodyBuilder();

            builder.HtmlBody = @"Thank you for using Safety Net as your UBI service!  Attached is your December 2020 invoice. 

<br/><br/>
Merry Christmas from Safety Net!

<br/><br/>
PS - Be sure to follow us on social media! - <a href=""https://twitter.com/SafetyNetUS"">Twitter</a>, <a href=""https://www.facebook.com/SafetyNetUS"">Facebook</a>, <a href=""https://www.instagram.com/safetynetus/"">Instagram</a>, <a href=""https://share.rizzle.tv/3VtBmL"">Rizzle</a>";

            builder.Attachments.Add(attachmentFilePath);

            mailMessage.Body = builder.ToMessageBody();

            await smtpClient.SendAsync(mailMessage);
        }

        private static void OpenPdf(string pathToPdf)
        {
            new Process
            {
                StartInfo = new ProcessStartInfo(pathToPdf)
                {
                    UseShellExecute = true
                }
            }.Start();
        }

        private static async Task<string> GenerateInvoiceAsync(string recipientEmail, string recipientId, int invoiceNumber, string invoiceDirectoryPath)
        {
            using var httpClient = new HttpClient();
            var invoiceFilePath = Path.Combine(invoiceDirectoryPath, $"202012-SafetyNetInvoice-{recipientId}.pdf");
            var values = new Dictionary<string, string>
                {
                    { "from", "Safety Net\nJerome Bell, Jr.\n\nhelp@safetynetus.io\nhttps://safetynetus.io" },
                    { "to", recipientEmail },
                    { "logo", "https://safetynetus.io/version-test/logo-1200x1200.png" },
                    { "number", $"{invoiceNumber}" },
                    { "date", "December 1, 2020" },
                    { "due_date", "December 1, 2020" },
                    { "items[0][name]", "Universal Basic Income\nYour Distribution of Universal Basic Income" },
                    { "items[0][quantity]", "1" },
                    { "items[0][unit_cost]", "-23.80" },
                    { "items[1][name]", "Service Fee\nStandard 0.1% Universal Basic Income distribution service fee" },
                    { "items[1][quantity]", "1" },
                    { "items[1][unit_cost]", "0.02" },
                    { "notes", "Thanks for being an awesome customer!" }
                };

            var content = new FormUrlEncodedContent(values);
            var response = await httpClient.PostAsync("https://invoice-generator.com", content);

            using var writer = new FileStream(invoiceFilePath, FileMode.OpenOrCreate);
            var stream = await response.Content.ReadAsStreamAsync();
            await stream.CopyToAsync(writer);
            writer.Close();

            return invoiceFilePath;
        }
    }
}
