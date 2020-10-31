using Common;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using MSALUtils;
using System;
using System.Security;
using System.Threading.Tasks;

namespace ANO.CSOM.DotNetCore
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //var Scopes = new string[] { "https://graph.microsoft.com/.default","https://microsoft.sharepoint-df.com/AllSites.Manage" };
            //var ClientId = CommonCredentials.ClientId;
            //var app = PublicClientApplicationBuilder.Create(ClientId)
            //    .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs).Build();
            //string username = CommonCredentials.UserName;
            //// Console.Write("Enter Password: ");
            //string pwd = CommonCredentials.Password;
            //SecureString password = new SecureString();
            //foreach (char c in pwd)
            //    password.AppendChar(c);

            //var App = new PublicAppUsingUsernamePassword(app);
            //var result = App.AcquireATokenFromCacheOrUsernamePasswordAsync(Scopes, username, password).GetAwaiter().GetResult();

            //var tkn = await App.AcquireATokenFromCacheOrUsernamePasswordAsync(Scopes, username, password);
            var tkn = GetAccessTokenUsingUsernamePassword().GetAwaiter().GetResult();
            //var publicClient = PublicClientApplicationBuilder.Create("");
            //new PublicClientApplication().AcquireTokenByUsernamePassword()
            ClientContext clientContext = new ClientContext("https://brgrp.sharepoint.com");
            clientContext.ExecutingWebRequest +=
                 delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + tkn.AccessToken;
                };

            //var clientContext = GetClientContextWithAccessToken(siteUrl, accessToken);

            Web web = clientContext.Web;

            clientContext.Load(web);

            clientContext.ExecuteQuery();

            Console.WriteLine(web.Title);
        }

        public static async Task<AuthenticationResult> GetAccessTokenUsingUsernamePassword()
        {
            var Scopes = new string[] { "https://brgrp.sharepoint.com/.default" };
            var ClientId = CommonCredentials.ClientId;
            var app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs).Build();
            string username = CommonCredentials.UserName;
            // Console.Write("Enter Password: ");
            string pwd = CommonCredentials.Password;
            SecureString password = new SecureString();
            foreach (char c in pwd)
                password.AppendChar(c);

            var App = new PublicAppUsingUsernamePassword(app);
            var result = App.AcquireATokenFromCacheOrUsernamePasswordAsync(Scopes, username, password).GetAwaiter().GetResult();

            var tkn = await App.AcquireATokenFromCacheOrUsernamePasswordAsync(Scopes, username, password);
            return tkn;

        }

    }
}
