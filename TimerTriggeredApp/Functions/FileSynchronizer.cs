using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using TimerTriggeredApp.Helper;

namespace TimerTriggeredApp.Functions
{
    public static class FileSynchronizer
    {
        [FunctionName("FileSynchronizer")]
        public static async Task Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            var userDisplayName = await GetUserInfo();
            log.LogInformation(userDisplayName);
        }

        public static async Task<string> GetUserInfo()
        {
            // Create a new instance of GraphServiceClient with the authentication provider.
            var graphClient = AuthProvider.getAuthenticatedGraphClient();
            var lists = await graphClient.Sites["Root"].Lists.Request().GetAsync();

            return lists[4].WebUrl;
        }
    }
}
