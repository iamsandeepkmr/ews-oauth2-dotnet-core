using System;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace EWS_OAuth2
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            // Using Microsoft.Identity.Client
            var cca = ConfidentialClientApplicationBuilder
                .Create("f7142365-b662-44c9-afec-6f3105a20143")      //client Id
                .WithClientSecret("z3g9fg-Y-15~imqr4GDQxC1M..55Zof0-~")
                .WithTenantId("50291d04-103e-4a6a-b699-e78122e371f5")
                .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            try
            {
                // Get token
                var authResult = await cca.AcquireTokenForClient(ewsScopes)
                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
                ewsClient.ImpersonatedUserId =
                    new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "test@testdomain.onmicrosoft.com");

                //Include x-anchormailbox header
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", "test@testdomain.onmicrosoft.com");

                // Make an EWS call to list folders on exhange online
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                foreach (var folder in folders.Result)
                {
                    Console.WriteLine($"Folder: {folder.DisplayName}");
                }

                // Make an EWS call to read 50 emails (last 5 days) from Inbox folder
                TimeSpan ts = new TimeSpan(-5, 0, 0, 0);
                DateTime date = DateTime.Now.Add(ts);
                SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                var findResults = ewsClient.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));
                foreach (var mailItem in findResults.Result)
                {
                    Console.WriteLine($"Subject: {mailItem.Subject}");
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            }
        }
    }
}
