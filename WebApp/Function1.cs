using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

namespace WebApp
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string element = req.GetQueryNameValuePairs()
            //string element = req.Content.ReadAsStringAsync())
                .FirstOrDefault(q => string.Compare(q.Key, "element", true) == 0).Value;
             


            if (element == null)
            {
                // Get request body
                dynamic data = await req.Content.ReadAsAsync<object>();
                element = data?.name;
            }

            return element == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, element);
        }
    }
}
