using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

using Azure.Connectors.MicrosoftTeams;
using Azure.Connectors.TextAnalytics;
using Azure.Connectors.TextAnalytics.Models;

using Microsoft.Rest.TransientFaultHandling;
using Microsoft.Rest.Azure;

namespace prbeegala.teamssentiment
{
    public static class TeamsSentimentScore
    {
        [FunctionName("TeamsSentimentScore")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Replace with your connection string for MicrosoftTeams
            var teamsConnector = MicrosoftTeamsConnector.Create("endpoint=https://47e705ee2f46a8cd.12.common.logic-westcentralus.azure-apihub.net/apim/teams/b30df2dac2084c59a04f560980816892;auth=managed");

            // Add resiliency to outbound calls
            // For retry policy documentation see 
            // https://docs.microsoft.com/en-us/dotnet/api/microsoft.rest.transientfaulthandling?view=azure-dotnet
            var retryStrategy = new FixedIntervalRetryStrategy(Int32.MaxValue, TimeSpan.FromTicks(1));
            var retryPolicy = new RetryPolicy(new HttpStatusCodeErrorDetectionStrategy(), retryStrategy);
            teamsConnector.SetRetryPolicy(retryPolicy);

            var teams = await teamsConnector.GetAllTeamsAsync();
            var team = teams.Value.FirstOrDefault(t => t.DisplayName.Equals("Kingfisher FY21"));
            var channels = await teamsConnector.GetChannelsForGroupAsync(team.Id);
            var channel = channels.Value.FirstOrDefault(c => c.DisplayName.Equals("ipaas RFP"));

            string lastMessage = string.Empty;
            try {
                var messages = await teamsConnector.GetMessagesFromChannelAsync(team.Id, channel.Id);
                lastMessage = messages.First().Body.Content;

                // Paging - Some of the operations might need paging 
                // Example in the above case
                //    if(messages.NextPageLink != null)
                //      var  moreMessages = await teamsConnector.GetMessagesFromChannelNextAsync(messages.NextPageLink);
            } catch(CloudException cloudException) {
               // provide below information when reporting runtime issues 
               // "x-ms-client-request-id", RequestUri and Method
               var requestUri = cloudException.Request.RequestUri.ToString();
               var clientRequestId = cloudException.Request.Headers["x-ms-client-request-id"].First();

               var statusCode = cloudException.Response.StatusCode.ToString();
               var reasonPhrase = cloudException.Response.ReasonPhrase;
               log.LogError("RequestUri '{0}' ClientRequestId '{1}' StatusCode '{2}' ReasonPhrase '{3}'", 
                    requestUri, clientRequestId, statusCode, reasonPhrase);

               throw cloudException;
            }
           
            // Replace with your connection string for TextAnalytics
            var cognitiveTextAnalyticsService = TextAnalyticsConnector.Create("endpoint=https://47e705ee2f46a8cd.12.common.logic-westcentralus.azure-apihub.net/apim/cognitiveservicestextanalytics/d655a1436b414876957f6638a3bd4276;auth=managed");
            cognitiveTextAnalyticsService.SetRetryPolicy(new RetryPolicy<HttpStatusCodeErrorDetectionStrategy>(3));

            var sentimentScore = await cognitiveTextAnalyticsService.Sentiment.DetectSentimentV2Async(new MultiLanguageInput { Language = "en", Text = lastMessage });

            return new OkObjectResult(sentimentScore);
        }
    }
}
