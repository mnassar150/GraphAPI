using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace EmailSenderApp
{
    class Program
    {
        private static string clientId = "915a5ce2-9986-4540-a5a7-7caa4378052e";
        private static string tenantId = "5d23882c-d9f0-4e2e-84a6-0f290d7fbdce";
        private static string clientSecret = "WASFADFADFASDFASDFeLFaMY";
        private static string userEmail = "mnassar365@outlook.com";

        static async Task Main(string[] args)
        {
            var confidentialClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                var authResult = await confidentialClient
                    .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                    .ExecuteAsync();

                requestMessage.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            }));

            var filePath = @"C:\Users\mnass\Downloads\pythonApp\Report.txt";
            var fileName = Path.GetFileName(filePath);
            var fileContent = await System.IO.File.ReadAllBytesAsync(filePath);



            var email = new Microsoft.Graph.Message
{
    Subject = "Your Requested Report from OnPrem",
    Body = new Microsoft.Graph.ItemBody
    {
        ContentType = Microsoft.Graph.BodyType.Text,
        Content = "Please find the attached report."
    },
    ToRecipients = new List<Microsoft.Graph.Recipient>
    {
        new Microsoft.Graph.Recipient
        {
            EmailAddress = new Microsoft.Graph.EmailAddress
            {
                Address = userEmail
            }
        }
    },
    Attachments = new MessageAttachmentsCollectionPage
    {
        new Microsoft.Graph.FileAttachment
        {
            Name = fileName,
            ContentType = "application/octet-stream",
            ContentBytes = fileContent
        }
    }
};


            try
            {
                await graphClient.Users["blablabla@1590.eu"]
                    .SendMail(email, true)
                    .Request()
                    .PostAsync();

                Console.WriteLine("Email sent successfully.");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }
    }
}
