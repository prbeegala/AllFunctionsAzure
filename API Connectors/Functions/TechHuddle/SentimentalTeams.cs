using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Azure.Connectors.MicrosoftTeams;
using System.Linq;
using Azure.Connectors.TextAnalytics;
using Azure.Connectors.TextAnalytics.Models;

namespace prbeegala.functions
{
    public static class SentimentalTeams
    {
        [FunctionName("SentimentalTeams")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //Connection string for Teams
            var teamsConnectionString = System.Environment.GetEnvironmentVariable("TEAMS_CONNECTION", EnvironmentVariableTarget.Process);
            var teamsConnector = MicrosoftTeamsConnector.Create(teamsConnectionString);

            //Read all my teams
            var teams = await teamsConnector.GetAllTeamsAsync();
            var team = teams.Value.FirstOrDefault (t => t.DisplayName.Equals("Azure Connectors Test"));

            //Read all the channels
            var channels = await teamsConnector.GetChannelsForGroupAsync(team.Id);
            var channel = channels.Value.FirstOrDefault(c => c.DisplayName.Equals("General"));

            //Read all the messages
            var messages = await teamsConnector.GetMessagesFromChannelAsync(team.Id, channel.Id);
            var lastMessage = string.Empty;
            lastMessage = messages.First().Body.Content;

            //Detect sentiment
            var textAnalyticsConnectionString = System.Environment.GetEnvironmentVariable("TEXTANALYTICS_CONNECTIONSTRING", EnvironmentVariableTarget.Process);
            var textAnalyticsConnector = TextAnalyticsConnector.Create(textAnalyticsConnectionString);
            var sentimentScore = await textAnalyticsConnector.Sentiment.DetectSentimentV2Async(new MultiLanguageInput{Language = "en", Text = lastMessage});

            return new OkObjectResult("The last message from teams is: "+ lastMessage+ " SENTIMENT SCORE IS: " + sentimentScore.Score);
        }
    }
}
