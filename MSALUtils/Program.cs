using Common;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace MSALUtils
{
    /// <summary>
    /// Reference From : https://github.com/Azure-Samples/active-directory-dotnetcore-console-up-v2
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var ClientId =CommonCredentials.ClientId;
            var app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs).Build();
            string username = CommonCredentials.UserName;
           // Console.Write("Enter Password: ");
            string pwd = CommonCredentials.Password;
            SecureString password = new SecureString();
            foreach (char c in pwd)
                password.AppendChar(c);
            var Scopes = new string[] { "https://graph.microsoft.com/.default" };
            Console.WriteLine("** Making Request to get GraphToken Using MSAL.NET ** \n");
            //var result =app.AcquireTokenByUsernamePassword(Scopes, username, password).ExecuteAsync().GetAwaiter().GetResult();

            #region Store MS GraphToken In Memory Caching With Username and Password flow
            var App = new PublicAppUsingUsernamePassword(app);
             var result = App.AcquireATokenFromCacheOrUsernamePasswordAsync(Scopes, username, password).GetAwaiter().GetResult();
            #endregion

            if (result != null)
            {
                Console.WriteLine("### RECEIVED TOKEN GraphToken Using MSAL.NET ###  \n  \n ");
                Console.WriteLine(result.AccessToken);
                var spClient = new HttpClient();
                spClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                spClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                spClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                var res = spClient.GetStringAsync("https://graph.microsoft.com/v1.0/me").Result;
                Console.WriteLine("\n** MS Graph Received Result **\n\n");
                Console.WriteLine(res);
            }
            else
            {
                Console.WriteLine("### NOT RECEIVED TOKEN ###");
            }
            Console.ReadKey();
        }
    }
}
