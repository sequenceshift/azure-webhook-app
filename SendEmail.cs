using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using SendGrid.Helpers.Mail;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;
namespace Company.Function

{
    public static class SendEmail
    {

        [FunctionName("SendEmailTimer")]
        [return: SendGrid(ApiKey = "SendGridApiKey")]
        public static async Task<SendGridMessage> Run([TimerTrigger("%MessageQueuerOccurence%")] TimerInfo myTimer,
          [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient webhookDocument,
          [CosmosDB("outDatabase", "UserCollection", ConnectionStringSetting = "CosmosDbConnectionString")] DocumentClient userDocument,
          ILogger log)
        {

            log.LogInformation($"SendEmailTimer executed at: {DateTime.Now}");
            DateTime nowDate = DateTime.Now;
            DateTime fromDate = nowDate.AddHours(-24);

            Uri collectionUri = UriFactory.CreateDocumentCollectionUri("outDatabase", "WebhookCollection");
            string q = "SELECT * FROM c WHERE c.payment_date_utc BETWEEN '" + fromDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + "'  AND '" + nowDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + "'";

            IDocumentQuery<dynamic> query = userDocument.CreateDocumentQuery(collectionUri, q,
                            new FeedOptions
                            {
                                PopulateQueryMetrics = true,
                                MaxItemCount = -1,
                                MaxDegreeOfParallelism = -1,
                                EnableCrossPartitionQuery = true
                            }

                            ).AsDocumentQuery();
            FeedResponse<dynamic> sqlResult = await query.ExecuteNextAsync();

            var msg = new SendGridMessage()
            {
                From = new EmailAddress(Environment.GetEnvironmentVariable("SenderEmail")),
                Subject = "Transactions Report from " + fromDate.ToString("yyyy-MM-dd_HH:mm") + " to " + nowDate.ToString("yyyy-MM-dd_HH:mm"),
                PlainTextContent = "Report-from-" + fromDate.ToString("yyyy-MM-dd_HH:mm:ss") + "-to-" + nowDate.ToString("yyyy-MM-dd_HH:mm:ss")

            };

            List<EmailAddress> emailList = new List<EmailAddress>();
            foreach (var email in Environment.GetEnvironmentVariable("RecipientEmail").Split(","))
            {
                emailList.Add(new EmailAddress(email.Trim()));
            }
            msg.AddTos(emailList);

            SendGrid.Helpers.Mail.Attachment att = new SendGrid.Helpers.Mail.Attachment
            {
                Content = Convert.ToBase64String(ReportHelper.CallGenerateExcelReport(sqlResult, userDocument, log)),
                Filename = "Report-from-" + fromDate.ToString("yyyy-MM-dd_HH:mm:ss") + "-to-" + nowDate.ToString("yyyy-MM-dd_HH:mm:ss") + ".xlsx",
                Type = "application/vnd.openxmlformats-officedocument.spreadsheet.excel",
                Disposition = "attachment"

            };
            msg.AddAttachment(att);

            return msg;
        }
    }
}