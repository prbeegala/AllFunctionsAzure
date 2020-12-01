using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Azure.Connectors.OneDrive;
using Azure.Connectors.Outlook;

namespace prbeegala.apiconnection
{
    public static class OneDriveFiles
    {
        [FunctionName("OneDriveFiles")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var outlookConnectionString = System.Environment.GetEnvironmentVariable("OUTLOOK_CONNECTION", EnvironmentVariableTarget.Process);
            var oneDriveConnectionString = System.Environment.GetEnvironmentVariable("ONEDRIVE_CONNECTION", EnvironmentVariableTarget.Process);

            var outlookConnector = OutlookConnector.Create(outlookConnectionString);
            var oneDriveConnector = OneDriveConnector.Create(oneDriveConnectionString);
//added comments
            var emails = await outlookConnector.Mail.GetEmailsV2Async(
                fetchOnlyUnread: true,
                fetchOnlyWithAttachment: true,
                includeAttachments: true
            );

            foreach(var email in emails.Value)
            {
                foreach(var attachemnt in email.Attachments)
                {
                    var bytes = attachemnt.ContentBytes;
                    Stream stream = new MemoryStream(bytes);

                    await oneDriveConnector.OneDriveFileData.CreateFileAsync("emailAttachments/", attachemnt.Name, stream);
                }
            }

            log.LogInformation("C# HTTP trigger function processed a request.");
            string responseMessage = "This HTTP triggered function executed successfully.";
            return new OkObjectResult(responseMessage);
        }
    }
}
