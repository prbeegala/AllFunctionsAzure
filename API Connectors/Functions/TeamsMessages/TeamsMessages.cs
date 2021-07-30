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

namespace prbeegala.functions
{
    public static class TeamsMessages
    {
        [FunctionName("TeamsMessages")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var teamsConnectionString = System.Environment.GetEnvironmentVariable("TEAMS_CONNECTION", EnvironmentVariableTarget.Process);
            var teamsConnector = MicrosoftTeamsConnector.Create(teamsConnectionString);
            
            var teams = await teamsConnector.GetAllTeamsAsync();
            var team = teams.Value.FirstOrDefault (t => t.DisplayName.Equals("Azure Connectors Test"));
            var channels = await teamsConnector.GetChannelsForGroupAsync(team.Id);
            var channel = channels.Value.FirstOrDefault(c => c.DisplayName.Equals("General"));
                       
            string lastMessage = string.Empty;
            try{
                var messages = await teamsConnector.GetMessagesFromChannelAsync(team.Id, channel.Id);
                lastMessage = messages.First().Body.Content;
            }
            catch (CloudException cloudException)
            {
                var requestUri = cloudException.Request.RequestUri.ToString();
                var clientRequestId = cloudException.Request.Headers["x-ms-client-request-id"].First();
                var statusCode = cloudException.Response.StatusCode.ToString();
                var reasonPhrase = cloudException.Response.ReasonPhrase; 
                log.LogError("RequestUri '{0}' ClientRequestId '{1}' StatusCode '{2}' ReasonPhrase '{3}'", 
                    requestUri, clientRequestId, statusCode, reasonPhrase);

               throw cloudException;                       
            }
            var cognitiveConnectionString = System.Environment.GetEnvironmentVariable("COGNITIVE_CONNECTION", EnvironmentVariableTarget.Process);
            var cognitiveConnector = TextAnalyticsConnector.Create(cognitiveConnectionString);
            var sentimentScore = await cognitiveConnector.Sentiment.DetectSentimentV2Async(new MultiLanguageInput { Language = "en", Text = lastMessage});
            return new OkObjectResult(lastMessage + sentimentScore.Score);
        }
    }
}
