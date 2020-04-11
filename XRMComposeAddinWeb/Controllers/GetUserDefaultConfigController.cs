using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Configuration;
using System.IdentityModel.Tokens;
using System.IO;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using XRMComposeAddinWeb.Models;
using System.Text;

namespace XRMComposeAddinWeb.Controllers
{
    public class GetUserDefaultConfigController : ApiController
    {
        // GET api/<controller>
        [HttpGet]
        public async Task<IHttpActionResult> Get(string useremail)
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

            return await GetUserDefaultConfig(useremail);
        }

        private async Task<IHttpActionResult> GetUserDefaultConfig(string useremail)
        {
            var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
            List<GetUserDefaultConfigInfo> userinfo = new List<GetUserDefaultConfigInfo>();
            var siteId = ConfigurationManager.AppSettings["ida:SiteId"];
            var defaultCongigId = ConfigurationManager.AppSettings["ida:DefaultConfigId"];
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

                // Initialize a Graph client
                GraphServiceClient graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                    // Add the Site Collection access token to each outgoing request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                            return Task.FromResult(0);
                        }));

                string filterString = "fields/UsersMail eq '" + useremail + "'";
                List<QueryOption> options = new List<QueryOption>()
                  {
                  new QueryOption("$expand","fields($select=id,Title,StatusID,UsersMail,CaseName,Category,CatName)"),
                  new QueryOption("$filter", filterString)
                  };

                var lcases = await graphClient.Sites[siteId].Lists[defaultCongigId].Items.Request(options).GetAsync();

                foreach (var lcase in lcases)
                {
                    userinfo.Add(new GetUserDefaultConfigInfo()
                    {
                        Title = lcase.Fields.AdditionalData["Title"].ToString(),
                        ID = lcase.Id,
                        StatusID = lcase.Fields.AdditionalData["StatusID"].ToString(),
                        UserMail = lcase.Fields.AdditionalData["UsersMail"].ToString(),
                        CaseName= lcase.Fields.AdditionalData["CaseName"].ToString(),
                       Category = lcase.Fields.AdditionalData["Category"].ToString(),
                       //Category = lcase.Fields.AdditionalData.Keys.Contains("Category") ? lcase.Fields.AdditionalData["Category"].ToString() : string.Empty,
                       CatName= lcase.Fields.AdditionalData["CatName"].ToString()
                    });
                }
            }
            return ResponseMessage(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonConvert.SerializeObject(userinfo, Formatting.Indented), Encoding.UTF8, "application/json") });


        }

    }
}
