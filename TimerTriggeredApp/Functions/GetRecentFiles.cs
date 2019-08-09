using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using TimerTriggeredApp.Helper;

namespace TimerTriggeredApp.Functions
{
    public static class GetRecentFiles
    {
        private static GraphServiceClient graphClient = null;

        [FunctionName("GetRecentFiles")]
        public static async Task Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            try
            {
                // Create a new instance of GraphServiceClient with the authentication provider.
                graphClient = AuthProvider.getAuthenticatedGraphClient();

                // Get recent files
                var recentFilesContent =await getRecentFiles();

                // Upload it as a file to SharePoint using Microsoft Graph
                await UploadReport(recentFilesContent);

                log.LogInformation("New report added to the Reports library in the root site collection");
            }
            catch(Exception ex)
            {
                log.LogError(ex.Message);
            }

            
        }

        public static async Task<string> getRecentFiles()
        {
            // For creating a CSV file to be uploaded in SharePoint
            var csv = new StringBuilder();

            // Add headers
            csv.AppendLine("File Name, Size (MB), Created Date");

            // Get Dev1 details
            var users = await graphClient.Users.Request().Filter("DisplayName eq 'Dev1'").Top(1).GetAsync();

            if (users.Count > 0)
            {
                var userId = users[0].Id;

                // Get 10 recent files
                var recentFiles = await graphClient.Users[userId].Drive.Recent().Request().Top(10).GetAsync();

                // Create an array of files
                var csvLines = (from file in recentFiles
                                select new object[]{
                                file.Name,
                                file.Size / 1024,
                                file.CreatedDateTime

                    }).ToList();

                // Append details to the CSV string builder
                csvLines.ForEach(line =>
                {
                    csv.AppendLine(string.Join(",", line));
                });
            }

            return csv.ToString();
        }

        public static async Task UploadReport(string csvContent)
        {
            // Get root site collection libraries
            var drives = await graphClient.Sites["Root"].Drives.Request().GetAsync();

            // Filter the libraries to get the Reports library
            var reportsLibrary = (from drive in drives where drive.Name=="Reports" select drive).FirstOrDefault();
            
            if (reportsLibrary != null)
            {
                byte[] byteArray = Encoding.ASCII.GetBytes(csvContent);

                // Create the file in the root folder
                var item = new DriveItem
                {
                    Name = $"Dev1RecentFiles-{DateTime.Now.ToString("yyyyMMddTHHmmss")}.csv",
                    File = new Microsoft.Graph.File { }
                };
                var addedItem = await graphClient.Sites["Root"].Drives[reportsLibrary.Id].Root.Children.Request().AddAsync(item);

                // Update the content
                using (var mStream = new MemoryStream(byteArray))
                {
                    var a = await graphClient.Sites["Root"].Drives[reportsLibrary.Id].Items[addedItem.Id].Content.Request().PutAsync<DriveItem>(mStream);
                }
            }
            
        }
    }
}
