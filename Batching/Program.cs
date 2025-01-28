using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Batching
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var tenantId = "tenant_id";
            var clientId = "client_id";
            var clientSecret = "client_secret";
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var client = new GraphServiceClient(credential);
            var userIds = await CreateUsers(client);

            var groupId = "group_id";
            await AddUsersToGroup(client, groupId, userIds);
        }

        static async Task<List<string>> CreateUsers(GraphServiceClient client)
        {
            var batchCollection = new BatchRequestContentCollection(client);

            var userIds = new List<string>();
            var requestIds = new List<string>();

            for (int i = 1; i <= 55; i++)
            {
                var newUser = new User
                {
                    AccountEnabled = true,
                    DisplayName = $"User {i}",
                    MailNickname = $"user{i}",
                    UserPrincipalName = $"user{i}@contoso.onmicrosoft.com",
                    PasswordProfile = new PasswordProfile
                    {
                        Password = $"Us3r-{i}",
                        ForceChangePasswordNextSignIn = false
                    }
                };

                var requestInfo = client.Users.ToPostRequestInformation(newUser);
                var requestId = $"user{i}";
                await batchCollection.AddBatchRequestStepAsync(requestInfo, requestId);
            }

            var batchResponse = await client.Batch.PostAsync(batchCollection);

            var responsesStatusCodes = await batchResponse.GetResponsesStatusCodesAsync();

            foreach (var response in responsesStatusCodes)
            {
                Console.WriteLine($"Response {response.Key} with status code {response.Value}");
                try
                {
                    var user = await batchResponse.GetResponseByIdAsync<User>(response.Key);
                    userIds.Add(user.Id);
                }
                catch (ServiceException error)
                {
                    var odataError = await KiotaJsonSerializer.DeserializeAsync<ODataError>(error.RawResponseBody);
                    Console.WriteLine($"Error: {odataError.Message}");
                }
            }

            return userIds;
        }

        static async Task AddUsersToGroup(GraphServiceClient client, string groupId, List<string> userIds)
        {
            var batchCollection = new BatchRequestContentCollection(client);

            for (int i = 0; i < userIds.Count; i += 20)
            {
                var selectedUserIds = userIds.Skip(i).Take(20);
                var members = selectedUserIds.Select(id => $"https://graph.microsoft.com/v1.0/directoryObjects/{id}");

                var requestBody = new Group
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "members@odata.bind" , members
                        }
                    }
                };

                var request = client.Groups[groupId].ToPatchRequestInformation(requestBody);

                await batchCollection.AddBatchRequestStepAsync(request, $"batch{i}");
            }

            var batchResponse = await client.Batch.PostAsync(batchCollection);

            var responsesStatusCodes = await batchResponse.GetResponsesStatusCodesAsync();

            foreach (var response in responsesStatusCodes)
            {
                Console.WriteLine($"Response {response.Key} with status code {response.Value}");
                if (response.Value != System.Net.HttpStatusCode.NoContent)
                {
                    try
                    {
                        var responseMessage = await batchResponse.GetResponseByIdAsync(response.Key);
                        var stringContent = await responseMessage.Content.ReadAsStringAsync();
                        var odataError = await KiotaJsonSerializer.DeserializeAsync<ODataError>(stringContent);
                        Console.WriteLine($"Error: {odataError.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }
            }
        }
    }
}
