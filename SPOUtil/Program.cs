using Common;
using InvokeSharePointRestAPI.Utils;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SPOUtil
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Credentials
            // Console.Write("Enter SharePoint Online Admin Site URL: ");
            string siteUrl = "https://brgrp-admin.sharepoint.com/";
            //Console.Write("Enter User Name: ");
            string userName = CommonCredentials.UserName;
            //  Console.Write("Enter Password: ");
            string password = CommonCredentials.Password;
            #endregion

            SecureString secureString = new SecureString();
            foreach (char c in password)
                secureString.AppendChar(c);

            var rootUri = new Uri(siteUrl);

            var credentials = new SharePointOnlineCredentials(userName, secureString);
            var cookie = credentials.GetAuthenticationCookie(rootUri);
            var reqUrl = $@"https://brgrp-admin.sharepoint.com/_vti_bin/client.svc/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$top=10&$filter=IsGroupConnected";
           
            //Use SharePoint REST API To Get Admin sites using CookieContainer
            var cc = new CookieContainer();
            cc.SetCookies(rootUri, cookie);

            Console.WriteLine("Requesting All site collection data");
            var allSites = CallSharePointOnlineAPI(reqUrl, HttpMethod.Get, null, null, null, cc).GetAwaiter().GetResult();


            Console.WriteLine("Received All site collection data");
            Console.WriteLine(allSites);


            #region Approach 2 Call REST API using CSOM
            var client = new SPHttpClient(rootUri, userName, password);
            var res = client.ExecuteJson(reqUrl); 
            #endregion

        }


        public static void GetUserProfiles(ClientContext clientContext)
        {
            // Replace the following placeholder values with the target SharePoint site and
            // target user.
            //   const string serverUrl = "http://serverName/";
            const string targetUser = "i:0#.f|membership|anomepani@brgrp.onmicrosoft.com";

            // Connect to the client context.
            //  ClientContext clientContext = new ClientContext(serverUrl);

            // Get the PeopleManager object and then get the target user's properties.
            PeopleManager peopleManager = new PeopleManager(clientContext);
            PersonProperties personProperties = peopleManager.GetPropertiesFor(targetUser);

            // Load the request and run it on the server.
            // This example requests only the AccountName and UserProfileProperties
            // properties of the personProperties object.
            clientContext.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
            clientContext.ExecuteQuery();

            foreach (var property in personProperties.UserProfileProperties)
            {
                Console.WriteLine(string.Format("{0}: {1}",
                    property.Key.ToString(), property.Value.ToString()));
            }
        }

        /// <summary>
        /// Get All ( Classic & Modern ) Site Collections from SharePoint Online Tenant
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        private static void getAllSiteCollections(string siteUrl, string userName, SecureString password)
        {


            var credentials = new SharePointOnlineCredentials(userName, password);
            ClientContext ctx = new ClientContext(siteUrl);
            ctx.Credentials = credentials;

            Tenant tenant = new Tenant(ctx);
            SPOSitePropertiesEnumerable siteProps = tenant.GetSitePropertiesFromSharePoint("0", true);
            ctx.Load(siteProps);
            ctx.ExecuteQuery();

            var authCookie = credentials.GetAuthenticationCookie(new Uri(siteUrl));

            Console.WriteLine("*************************************************");
            Console.WriteLine("Total Site Collections: " + siteProps.Count.ToString());
            foreach (var site in siteProps)
            {
                Console.WriteLine("{0} - {1}", site.Title, site.Template.ToString());
            }
        }

        /// <summary>
        /// Get Classic Site Collections from SharePoint Online Tenant
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        private static void getClassicSiteCollections(string siteUrl, string userName, SecureString password)
        {
            var credentials = new SharePointOnlineCredentials(userName, password);
            ClientContext ctx = new ClientContext(siteUrl);
            ctx.Credentials = credentials;

            //ctx.ExecutingWebRequest
            Tenant tenant = new Tenant(ctx);

            //var ac = ctx.GetAccessToken();
            //  ctx.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;
            var authCookie = credentials.GetAuthenticationCookie(new Uri(siteUrl));

            var headers = new Dictionary<string, string>()
                   {
                       {"Cookie",authCookie}
                   };

            var allSites = CallSharePointOnlineAPI($@"https://brgrp-admin.sharepoint.com/_vti_bin/client.svc/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$top=10&$filter=IsGroupConnected", HttpMethod.Get, null, headers, null, null).GetAwaiter().GetResult();

            SPOSitePropertiesEnumerable siteProps = tenant.GetSitePropertiesFromSharePoint("0", true);
            //tenant.LegacyAuthProtocolsEnabled = false;
            ctx.Load(siteProps);
            ctx.ExecuteQuery();
            Console.WriteLine("*************************************************");
            Console.WriteLine("Total Classic Collections: " + siteProps.Count.ToString());
            foreach (var site in siteProps)
            {
                Console.WriteLine("{0} - {1}", site.Title, site.Template.ToString());
            }
        }

        private static SecureString GetSecureString()
        {
            string password = "";
            SecureString securePassword = new SecureString();

            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        password = password.Substring(0, password.Length - 1);
                        int pos = Console.CursorLeft;
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            Console.WriteLine();



            //Convert string to secure string  
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();

            return securePassword;
        }

        private static async Task<string> CallSharePointOnlineAPI(string requestUrl, HttpMethod method, string body, Dictionary<string, string> addHeaders, string office365Token, CookieContainer cc)
        {
            string result = "";
            var handler = new HttpClientHandler() { CookieContainer = cc };
            using (var httpClient = new HttpClient(handler))
            {
                //httpClient.DefaultRequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                //httpClient.DefaultRequestHeaders.Add("User-Agent", @"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.94 Safari/537.36");
                HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //   request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", office365Token);

                if (addHeaders != null)
                {
                    foreach (KeyValuePair<string, string> item in addHeaders)
                    {
                        request.Headers.Add(item.Key, item.Value);
                    }
                }
                if (!string.IsNullOrEmpty(body))
                {
                    request.Content = new StringContent(body, Encoding.UTF8, "application/json");
                    request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                }

                using (HttpResponseMessage response = await httpClient.SendAsync(request))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        using (HttpContent content = response.Content)
                        {
                            result = await content.ReadAsStringAsync();
                        }
                    }
                    else
                    {
                        using (HttpContent content = response.Content)
                        {
                            var errorString = await content.ReadAsStringAsync();
                            throw new Exception(errorString);
                        }
                    }
                }
            }
            return result;
        }

    }
}
