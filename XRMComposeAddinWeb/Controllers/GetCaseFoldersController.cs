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
using Newtonsoft.Json.Linq;

namespace XRMComposeAddinWeb.Controllers
{
    public class GetCaseFoldersController : ApiController
    {
        private HttpClient _saveHttpClient = new HttpClient();
        [HttpPost]
        public async Task<IHttpActionResult> Post([FromBody]CaseInfo request)
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

            return await GetDocumentLibraryFolders(request);
        }

        private async Task<IHttpActionResult> GetDocumentLibraryFolders(CaseInfo driveinfo)
        {
            var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
            string driveid = ConfigurationManager.AppSettings["ida:CaseDriveId"];
            List<FolderInfo> folders = new List<FolderInfo>();
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

                //new QueryOption("search","contentclass:STS_Site")
                List<QueryOption> options = new List<QueryOption>()
                {
                    new QueryOption("$filter","folder ne null"),
                    new QueryOption("$select","id,name,webUrl")
                };

                //var libraryfolders = await graphClient.Drives[driveid].Root.Children.Request(options).GetAsync();
                //string folderName = MakeFileNameValid(string.Format("{0}-{1}", driveinfo.Title, driveinfo.ID));
                string caseFolderName = driveinfo.CaseFolderName;
                if (driveinfo.Level == "1")
                {
                    caseFolderName = await GetCaseFolderSharepoint(driveinfo.ID);
                    //var casefolders = await graphClient.Drives[driveid].Root.Children.Request(options).GetAsync();
                    //caseFolderName = GetCaseFolderName(casefolders, driveinfo.ID);
                }

                IDriveItemChildrenCollectionPage libraryfolders;
                if (driveinfo.Level == "1")
                {
                    libraryfolders = await graphClient.Drives[driveid].Root.ItemWithPath(caseFolderName).Children.Request(options).GetAsync();
                }
                else
                {
                    libraryfolders = await graphClient.Drives[driveid].Root.ItemWithPath(driveinfo.FolderPath).Children.Request(options).GetAsync();
                }


                foreach (var folder in libraryfolders)
                {
                    folders.Add(new FolderInfo()
                    {
                        Id = folder.Id,
                        Name = folder.Name,
                        WebUrl = folder.WebUrl,
                        CaseFolderName=caseFolderName
                    });
                }
            }

            //return Ok();
            return ResponseMessage(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonConvert.SerializeObject(folders, Formatting.Indented), Encoding.UTF8, "application/json") });
        }

        private string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).Replace("&", string.Empty).Replace(" ", string.Empty);
        }

        private string GetCaseFolderName(IDriveItemChildrenCollectionPage caseFolders, string itemid)
        {
            string caseFolderName = "";
            foreach(var folder in caseFolders)
            {
                string[] x = folder.Name.Split('-');
                if (x[x.Length - 1] == itemid)
                {
                    caseFolderName = folder.Name;
                    return caseFolderName;
                }
            }

            return caseFolderName;
        }

        private async Task<string> GetCaseFolderSharepoint(string itemid)
        {
            //string siteUrl = "https://ankerhan.sharepoint.com/sites/Sager";
            //OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();
            //using (var cc = authMgr.GetAppOnlyAuthenticatedContext(siteUrl, "50cd26dc-9a11-45cc-a6b1-90cbec596642", "fnpneN2ha5UPVRqIUGn6a0rukztKtPMOyJX2dLnt2kI"))
            //{
            //    cc.Load(cc.Web, p => p.Title);
            //    cc.ExecuteQuery();
            //    Console.WriteLine(cc.Web.Title);
            //};
            var furl = "https://prod-56.westeurope.logic.azure.com:443/workflows/7b826ecd39634e89b1b22cef54db749f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4Umz1p7kzWD3Y0KyMSkIFPTJyncSpN1k3WsF31lO8gw";
            sagerItemid fdata = new sagerItemid()
            {
                Itemid = itemid
            };
            HttpRequestMessage requestMsg = new HttpRequestMessage(new HttpMethod("POST"), furl);
            requestMsg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            requestMsg.Content = new StringContent(JsonConvert.SerializeObject(fdata), Encoding.UTF8, "application/json");
            HttpResponseMessage response = _saveHttpClient.SendAsync(requestMsg).Result;
            var content = await response.Content.ReadAsStringAsync();
            var json = JObject.Parse(content);
            var value = json["value"].ToList();
            string foldername = string.Empty;
            foreach (var item in value)
            {
                foldername = item["{Name}"].ToString();
            }
            return foldername;
        }
    }

    public class sagerItemid
    {
        [JsonProperty(PropertyName = "itemid")]
        public string Itemid { get; set; }
    }
}
