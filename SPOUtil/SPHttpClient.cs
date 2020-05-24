using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;

namespace InvokeSharePointRestAPI.Utils
{
    /// <summary>
    /// Http client for SharePoint Online
    /// Reference from : https://blog.vgrem.com/2015/04/04/consume-sharepoint-online-rest-service-using-net/
    /// </summary>
    public class SPHttpClient : HttpClient
    {
        public SPHttpClient(Uri webUri, string userName, string password) : base(new SPHttpClientHandler(webUri, userName, password))
        {
            BaseAddress = webUri;
        }

        //public object BaseAddress { get; private set; }


        /// <summary>
        /// Execure request method
        /// </summary>
        /// <param name="requestUri"></param>
        /// <param name="method"></param>
        /// <param name="headers"></param>
        /// <param name="payload"></param>
        /// <returns></returns>
        public JObject ExecuteJson<T>(string requestUri, HttpMethod method, IDictionary<string, string> headers, T payload)
        {
            HttpResponseMessage response;
            switch (method.Method)
            {
                case "POST":
                    var requestContent = new StringContent(JsonConvert.SerializeObject(payload));
                    requestContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                    DefaultRequestHeaders.Add("X-RequestDigest", RequestFormDigest());
                    if (headers != null)
                    {
                        foreach (var header in headers)
                        {
                            DefaultRequestHeaders.Add(header.Key, header.Value);
                        }
                    }
                    response = PostAsync(requestUri, requestContent).Result;
                    break;
                case "GET":
                    response = GetAsync(requestUri).Result;
                    break;
                default:
                    throw new NotSupportedException(string.Format("Method {0} is not supported", method.Method));
            }

            response.EnsureSuccessStatusCode();
            var responseContent = response.Content.ReadAsStringAsync().Result;
            return String.IsNullOrEmpty(responseContent) ? new JObject() : JObject.Parse(responseContent);
        }

        public JObject ExecuteJson<T>(string requestUri, HttpMethod method, T payload)
        {
            return ExecuteJson(requestUri, method, null, payload);
        }

        public JObject ExecuteJson(string requestUri)
        {
            return ExecuteJson(requestUri, HttpMethod.Get, null, default(string));
        }


        /// <summary>
        /// Request Form Digest
        /// </summary>
        /// <returns></returns>
        public string RequestFormDigest()
        {
            var endpointUrl = string.Format("{0}/_api/contextinfo", BaseAddress);
            var result = this.PostAsync(endpointUrl, new StringContent(string.Empty)).Result;
            result.EnsureSuccessStatusCode();
            var content = result.Content.ReadAsStringAsync().Result;
            var contentJson = JObject.Parse(content);
            return contentJson["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
        }
    }

}
