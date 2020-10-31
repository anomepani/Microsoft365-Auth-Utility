using Common;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ADALUtils
{
    /// <summary>
    /// Reference from : https://github.com/Azure-Samples/active-directory-dotnet-native-headless
    /// </summary>
    class Program
    {
        #region Init
        private static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static string tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);
        //
        // To authenticate to the To Do list service, the client needs to know the service's App ID URI.
        // To contact the To Do list service we need its URL as well.
        //
        private static string todoListResourceId = ConfigurationManager.AppSettings["todo:ResourceId"];
        //private static string todoListBaseAddress = ConfigurationManager.AppSettings["todo:TodoListBaseAddress"];
        private static AuthenticationContext authContext = null;
        #endregion
        static void Main(string[] args)
        {
            //100% Working Demo with Latest ADAL.NET Using Username and Password To Get Microsoft Graph Token
            // Tested with  <package id="Microsoft.IdentityModel.Clients.ActiveDirectory" version="5.2.2" targetFramework="net461" />
            //Configure authContext For Multi Tenancy with TokenCache in ADAL.NET

            authContext = new AuthenticationContext(authority, new FileCache());

            //authContext = new AuthenticationContext(authority, new TokenCache());

            //authContext = new AuthenticationContext(authority);

            #region Obtain token

            AuthenticationResult result = TryFetchTokenSilently().GetAwaiter().GetResult();

            if (result == null)
            {
                // Authenticate using Username and Password
                UserPasswordCredential uc = GetUserCredential();
              

                try
                {
                    result = authContext.AcquireTokenAsync(todoListResourceId, clientId, uc).Result;
                }
                catch (Exception ee)
                {
                    Console.WriteLine(ee);
                }
            }
            if (result != null)
            {
                Console.WriteLine("\n** MS Graph Token Received Using ADAL.NET **\n\n");
                Console.WriteLine(result.AccessToken);
            }
            var spClient = new HttpClient();
            spClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //spClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            spClient.DefaultRequestHeaders.Authorization= new AuthenticationHeaderValue("Bearer", result.AccessToken);
            var res = spClient.GetStringAsync("https://graph.microsoft.com/v1.0/me").Result;
            Console.WriteLine("\n** MS Graph Received Result **\n\n");

            res = spClient.GetStringAsync("https://graph.microsoft.com/v1.0/users").Result;
            Console.WriteLine("\n** MS Graph users Received Result **\n\n");
            Console.WriteLine(res);
            #endregion

        }

        /// <summary>
        ///  Gather user credentials form the command line
        /// </summary>
        /// <returns></returns>
        static UserPasswordCredential GetUserCredential()
        {
            string user = CommonCredentials.UserName;
            string password = CommonCredentials.Password;
            return new UserPasswordCredential(user, password);
        }

        /// <summary>
        /// Fetch Token Silently from FileCache
        /// </summary>
        /// <returns></returns>
        private static async Task<AuthenticationResult> TryFetchTokenSilently()
        {
            AuthenticationResult result = null;

            // first, try to get a token silently
            try
            {
                result = await authContext.AcquireTokenSilentAsync(todoListResourceId, clientId);
                Console.WriteLine("token from  the cache");
                return result;
            }
            catch (AdalException adalException)
            {
                // There is no token in the cache; prompt the user to sign-in.
                if (adalException.ErrorCode == AdalError.FailedToAcquireTokenSilently
                    || adalException.ErrorCode == AdalError.InteractionRequired)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("No token in the cache");
                    return result;
                }

                // An unexpected error occurred.
            }

            return result;
        }

        // Empties the token cache
        static void ClearCache()
        {
            authContext.TokenCache.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Token cache cleared.");
        }

    }
}
