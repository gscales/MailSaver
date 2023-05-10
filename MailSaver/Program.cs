
using System.Security;
using Microsoft.Graph;
using System.Net;
using Azure.Identity;
Console.WriteLine("Hello, World Starting MailSaver");
string user = "gscales@.....com";
string saveDirectory = "c:\\temp\\";
var graphclient = await GetGraphClientWithClientCredentials("...", await EncodeSecret("..."), "...");
var profile = await graphclient.Users[user].GetAsync();
Console.WriteLine($"Hello {profile?.DisplayName}");
Console.WriteLine($"Let wait for you new messages and save them to the directory {saveDirectory}");
await TrackNewMail(graphclient, user, saveDirectory);

static async Task<bool> TrackNewMail(GraphServiceClient graphServiceClient, string userId, string saveDirectory)
{
    string filterString = "receivedDateTime ge " + DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture);
    int pollInterval = 30;
    var requestInfo = graphServiceClient.Users[userId].MailFolders["inbox"].Messages.Delta.ToGetRequestInformation();
    requestInfo.UrlTemplate = requestInfo.UrlTemplate?.Insert(requestInfo.UrlTemplate.Length - 1, ",changeType");
    requestInfo.QueryParameters.Add("%24filter", filterString);
    requestInfo.QueryParameters.Add("changeType", "created");
    var messagesDeltas = graphServiceClient.RequestAdapter.SendAsync<Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaResponse>(requestInfo, Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaResponse.CreateFromDiscriminatorValue).GetAwaiter().GetResult();
    while (messagesDeltas != null)
    {
        if (messagesDeltas.Value == null || messagesDeltas.Value.Count <= 0)
        {
            Console.WriteLine("No changes...");
        }
        else
        {
            var morePagesAvailable = false;

            do
            {
                if (messagesDeltas == null || messagesDeltas.Value == null)
                {
                    continue;
                }

                // Process current page
                foreach (var message in messagesDeltas.Value)
                {
                    string fileName = saveDirectory + "\\" + Guid.NewGuid().ToString() + ".eml";
                    Console.WriteLine(message.Subject);
                    Console.WriteLine($"Saving to {fileName}");
                    var mimeStream = await graphServiceClient.Users[userId].Messages[message.Id].Content.GetAsync();
                    if(mimeStream != null)
                    {
                        using (var fileStream = File.Create(fileName))
                        {
                            mimeStream.Seek(0, SeekOrigin.Begin);
                            mimeStream.CopyTo(fileStream);
                            mimeStream.Close();
                        }
                    }

                }

                morePagesAvailable = !string.IsNullOrEmpty(messagesDeltas.OdataNextLink);

                if (morePagesAvailable)
                {
                    // If there is a OdataNextLink, there are more pages
                    // Get the next page of results
                    var request = new Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaRequestBuilder(messagesDeltas.OdataNextLink, graphServiceClient.RequestAdapter);
                    messagesDeltas = request.GetAsync().GetAwaiter().GetResult();
                }
            }
            while (morePagesAvailable);
        }

        Console.WriteLine($"Processed current delta. Will check back in {pollInterval} seconds.");

        // Once we've iterated through all of the pages, there should
        // be a delta link, which is used to request all changes since our last query
        var deltaLink = messagesDeltas?.OdataDeltaLink;
        if (!string.IsNullOrEmpty(deltaLink))
        {
            Task.Delay(pollInterval * 1000).GetAwaiter().GetResult();
            var request = new Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaRequestBuilder(deltaLink, graphServiceClient.RequestAdapter);
            messagesDeltas = request.GetAsync().GetAwaiter().GetResult();
        }
        else
        {
            Console.WriteLine("No @odata.deltaLink found in response!");
        }
    }
    return true;
}

static async Task<GraphServiceClient> GetGraphClientWithClientCredentials(string clientId, SecureString clientSecret, string tennatId)
{
    ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tennatId, clientId, await DecodeSecret(clientSecret));
    return new GraphServiceClient(clientSecretCredential);
}

static async Task<string> DecodeSecret(SecureString clientSecret)
{
    return new NetworkCredential("", clientSecret).Password;
}
static async Task<SecureString> EncodeSecret(string clientSecret)
{
    return new NetworkCredential("", clientSecret).SecurePassword;
}
