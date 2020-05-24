using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace InvokeSharePointRestAPI.Utils
{
    /// <summary>
    /// Custom Http Handler for Sharepoint Rest
    /// Reference from : https://blog.vgrem.com/2015/04/04/consume-sharepoint-online-rest-service-using-net/
    /// </summary>
    public class SPHttpClientHandler : HttpClientHandler
    {

        public SPHttpClientHandler(Uri webUri, string userName, string password)
        {
            CookieContainer = GetAuthCookies(webUri, userName, password);
            FormatType = FormatType.JsonVerbose;
        }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            if (FormatType == FormatType.JsonVerbose)
            {
                ////request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json;odata=verbose"));
                //request.Headers.Add("Accept", "application/json;odata=verbose");
                request.Headers.Add("Accept", "application/json;odata=nometadata");
            }
            return base.SendAsync(request, cancellationToken);
        }


        /// <summary>
        /// Retrieve SPO Auth Cookies 
        /// </summary>
        /// <param name="webUri"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private static CookieContainer GetAuthCookies(Uri webUri, string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(userName, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(webUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(webUri, authCookie);
            return cookieContainer;
        }


        public FormatType FormatType { get; set; }
    }
    public enum FormatType
    {
        JsonVerbose,
        Xml
    }
}
