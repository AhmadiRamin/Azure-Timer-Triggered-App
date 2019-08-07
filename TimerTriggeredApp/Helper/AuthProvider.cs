using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using TimerTriggeredApp.Constant;

namespace TimerTriggeredApp.Helper
{
    public static class AuthProvider
    {
        public static IConfidentialClientApplication GetClientApplication()
        {
            string clientId = Environment.GetEnvironmentVariable(Configs.ClientIdKey);
            string tenantId = Environment.GetEnvironmentVariable(Configs.TenantIDKey);
            string secretKey = Environment.GetEnvironmentVariable(Configs.SecretKey);

            if (string.IsNullOrEmpty(clientId))
                throw new ArgumentException($"Missing required environment variable '{Configs.ClientIdKey}'");

            if (string.IsNullOrEmpty(tenantId))
                throw new ArgumentException($"Missing required environment variable '{Configs.TenantIDKey}'");

            if (string.IsNullOrEmpty(secretKey))
                throw new ArgumentException($"Missing required environment variable '{Configs.SecretKey}'");

            return ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(secretKey)
            .WithTenantId(tenantId)
            .Build();
        }

        public static GraphServiceClient getAuthenticatedGraphClient()
        {
            ClientCredentialProvider authProvider = new ClientCredentialProvider(GetClientApplication());
            return new GraphServiceClient(authProvider);
        }
    }
}
