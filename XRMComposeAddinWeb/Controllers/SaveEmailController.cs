using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.IdentityModel.Tokens;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using XRMComposeAddinWeb.Models;

namespace XRMComposeAddinWeb.Controllers
{
    public class SaveEmailController : ApiController
    {
        private HttpClient _saveHttpClient = new HttpClient();
        [HttpPost]
        public async Task<IHttpActionResult> Post([FromBody]SaveEmailRequest request)
        {
            if (Request.Headers.Contains("Authorization"))
            {
                // Request contains bearer token, validate it
                var scopeClaim = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope");
                if (scopeClaim != null)
                {
                    // Check the allowed scopes
                    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
                    if (!addinScopes.Contains("access_as_user"))
                    {
                        return BadRequest("The bearer token is missing the required scope.");
                    }
                }
                else
                {
                    return BadRequest("The bearer token is invalid.");
                }

                var issuerClaim = ClaimsPrincipal.Current.FindFirst("iss");
                var tenantIdClaim = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid");
                if (issuerClaim != null && tenantIdClaim != null)
                {
                    // validate the issuer
                    string expectedIssuer = string.Format("https://login.microsoftonline.com/{0}/v2.0", tenantIdClaim.Value);
                    if (string.Compare(issuerClaim.Value, expectedIssuer, StringComparison.OrdinalIgnoreCase) != 0)
                    {
                        return BadRequest("The token issuer is invalid.");
                    }
                }
                else
                {
                    return BadRequest("The bearer token is invalid.");
                }
            }
            else
            {
                return BadRequest("Authorization is not valid");
            }

            return await SaveEmail(request);
        }

        private async Task<IHttpActionResult> SaveEmail(SaveEmailRequest request)
        {
            string siteid = ConfigurationManager.AppSettings["ida:SiteId"];
            string listId= ConfigurationManager.AppSettings["ida:OutlookEmailListId"];
            var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
            if (bootstrapContext != null)
            {
                // Use MSAL to invoke the on-behalf-of flow to exchange token for a Graph token
                UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
                ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:AppPassword"]);
                ConfidentialClientApplication cca = new ConfidentialClientApplication(
                    ConfigurationManager.AppSettings["ida:AppId"],
                    ConfigurationManager.AppSettings["ida:RedirectUri"],
                    clientCred, null, null);

                string[] graphScopes = { "Files.ReadWrite", "Mail.Read", "Sites.ReadWrite.All" };

                AuthenticationResult authResult = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion);

                //// Initialize a Graph client
                //GraphServiceClient graphClient = new GraphServiceClient(
                //    new DelegateAuthenticationProvider(
                //        (requestMessage) =>
                //        {
                //            // Add the Site Collection access token to each outgoing request
                //            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                //            return Task.FromResult(0);
                //        }));

                string requestUrl = string.Format("https://graph.microsoft.com/v1.0/sites/{0}/lists('{1}')/items", siteid, listId);
                SaveEmailPost data = new SaveEmailPost()
                {
                    fields = request
                };
                var posdata = JsonConvert.SerializeObject(data);
                HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("POST"), requestUrl);
                requestMsg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                requestMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                requestMsg.Content = new StringContent(posdata, Encoding.UTF8, "application/json");
                HttpResponseMessage response = _saveHttpClient.SendAsync(requestMsg).Result;
                if (response.IsSuccessStatusCode)
                {
                    return Ok();
                }
                else
                {
                    return BadRequest("Error while saving the item. Please contact the administrator.");
                }
            }
            else
            {
                return BadRequest("Error while fetching the BootstrapContext. Please contact the administrator.");
            }
        }
    }
}
