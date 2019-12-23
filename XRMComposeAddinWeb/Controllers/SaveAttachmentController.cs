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
    public class SaveAttachmentController : ApiController
    {
        [HttpPost]
        public async Task<IHttpActionResult> Post([FromBody] SaveAttachmentRequest request)
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

            return await GetAttachments(request);
        }

        private async Task<IHttpActionResult> GetAttachments(SaveAttachmentRequest request)
        {
            string driveid = ConfigurationManager.AppSettings["ida:CaseDriveId"];
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

                // Initialize a Graph client
                GraphServiceClient graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            // Add the Site Collection access token to each outgoing request
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                            return Task.FromResult(0);
                        }));

                foreach (string attachmentId in request.attachmentIds)
                {
                    var attachment = await graphClient.Me.Messages[request.messageId].Attachments[attachmentId].Request().GetAsync() as FileAttachment;
                    string caseFolderName = request.caseFolderName;
                    if (attachment.IsInline == false)
                    {
                      // MemoryStream fileStream = new MemoryStream(attachment.ContentBytes);
                        if (attachment.Size < (4 * 1024 * 1024))
                        {
                            using (MemoryStream fileStream = new MemoryStream(attachment.ContentBytes))
                            {
                               // string caseFolderName = MakeFileNameValid(request.caseFolderName);
                                bool success = await SaveFileToSharePoint(graphClient, attachment.Name, fileStream, driveid, request.folderName, caseFolderName);
                                if (!success)
                                {
                                    return BadRequest("Failed to upload the file to the Sharepoint document Library");
                                }
                            }
                        }
                        else if (attachment.Size > (4 * 1024 * 1024))
                        {
                            try
                            {
                                //  List<ResultsItem> items = new List<ResultsItem>();
                                // Create the upload session. The access token is no longer required as you have session established for the upload.  
                                //string caseFolderName = MakeFileNameValid(request.caseFolderName);
                               bool success= await SaveLargeFileToSharePoint(graphClient,attachment.ContentBytes ,attachment.Name,driveid, request.folderName, caseFolderName);
                                if (!success)
                                {
                                    return BadRequest("Failed to upload the file to the Sharepoint document Library");
                                }
                            }
                            catch (Exception ex) {  }

                        }

                        else
                        {
                            // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createuploadsession
                            // and
                            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Models/FilesService.cs
                            return BadRequest("Attachments with size >4MB are restricted from uploading");
                        }
                    }
                    //else
                    //{
                    //    return BadRequest("There is no attachment found in the mail.Please unselect the attachment option");
                    //}

                }
            }
            return Ok();
        }

        private async Task<bool> SaveFileToSharePoint(GraphServiceClient client, string fileName, Stream fileContent, string driveId, string foldername, string caseFolderName)
        {
            try
            {
                if (!string.IsNullOrEmpty(foldername))
                {
                    string path = $"{caseFolderName}/{foldername}/{MakeFileNameValid(fileName)}";
                    DriveItem item = await client.Drives[driveId].Root.ItemWithPath(path).Content.Request().PutAsync<DriveItem>(fileContent);
                }
                else
                {
                    string path = $"{caseFolderName}/{MakeFileNameValid(fileName)}";
                    DriveItem item = await client.Drives[driveId].Root.ItemWithPath(path).Content.Request().PutAsync<DriveItem>(fileContent);
                }

            }
            catch (ServiceException)
            {
                return false;
            }

            return true;
        }
        private async Task<bool> SaveLargeFileToSharePoint(GraphServiceClient client,byte[] contentBytes, string fileName, string driveId, string foldername, string caseFolderName)
        {
            using (MemoryStream fileStream = new MemoryStream(contentBytes))
            {
                try
                {
                    string path = string.Empty;
                   
                    if (!string.IsNullOrEmpty(foldername))
                    {
                         path = $"{caseFolderName}/{foldername}/{MakeFileNameValid(fileName)}";
                    }
                    else
                    {
                         path = $"{caseFolderName}/{MakeFileNameValid(fileName)}";
                    }
                    UploadSession uploadSession = await client.Drives[driveId].Root.ItemWithPath(path).CreateUploadSession().Request().PostAsync();

                    int maxChunkSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
                    ChunkedUploadProvider provider = new ChunkedUploadProvider(uploadSession, client, fileStream, maxChunkSize);

                    // Set up the chunk request necessities.
                    IEnumerable<UploadChunkRequest> chunkRequests = provider.GetUploadChunkRequests();
                    byte[] readBuffer = new byte[maxChunkSize];
                    List<Exception> trackedExceptions = new List<Exception>();
                    DriveItem uploadedFile = null;

                    // Upload the chunks.
                    foreach (var request1 in chunkRequests)
                    {
                        // Do your updates here: update progress bar, etc.
                        // ...
                        // Send chunk request
                        UploadChunkResult result = await provider.GetChunkRequestResponseAsync(request1, readBuffer, trackedExceptions);

                        if (result.UploadSucceeded)
                        {
                            uploadedFile = result.ItemResponse;

                        }
                    }


                    // Check that upload succeeded.
                    if (uploadedFile == null)
                    {
                        // Retry the upload
                        // ...
                    }
                    return (uploadedFile != null);
                }
                catch (ServiceException)
                {
                    return false;
                }
            }
        }
    
        private string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).Replace("&", string.Empty).Replace(" ", string.Empty);
        }
    }
}
