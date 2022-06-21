using System;

using System.IO;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using SendGrid.Helpers.Mail;
using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
namespace Company.Function
{
    public class SendEmail
    {
        public byte[] excelAttachment(IEnumerable<object> inputDocument)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            MemoryStream memorystream = new MemoryStream();
            using (ExcelPackage excel = new ExcelPackage())
            {

                var worksheet = excel.Workbook.Worksheets.Add("Report1");

                // build exel header
                var headerRow = new List<string[]>();
                var arr = Enum.GetNames(typeof(ReportHeader));
                headerRow.Add(arr);
                String upRange = (headerRow[0].Length + 64) > 90 ? "A" + Char.ConvertFromUtf32(headerRow[0].Length + 38) : Char.ConvertFromUtf32(headerRow[0].Length + 64);
                string headerRange = "A1:" + upRange + "1";
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                // build report content
                int row = 2;
                int numCol = worksheet.Dimension.Columns;
                foreach (var documentItem in inputDocument)
                {
                    var json = JsonConvert.SerializeObject(documentItem);
                    JToken responseBody = JToken.FromObject(documentItem);
                    foreach (ReportHeader a in Enum.GetValues(typeof(ReportHeader)))
                    {
                        string cellValue = responseBody[a.ToString()] != null ? responseBody[a.ToString()].ToString().Replace("\n", "").Replace("\r", "") : null;
                        worksheet.Cells[row, (int)a + 1].Value = cellValue;
                    }
                    row = row + 1;
                }

                excel.SaveAs(memorystream);

            }
            return memorystream.ToArray();
        }


    
        const string query = "SELECT * FROM c WHERE c.payment_date BETWEEN DateTimeAdd(\"hh\", -24, GetCurrentDateTime())  AND GetCurrentDateTime()";
      
        [FunctionName("SendEmailTimer")]
        [return: SendGrid(ApiKey = "SendGridApiKey")]
        public SendGridMessage Run([TimerTrigger("*/15 * * * * *")] TimerInfo myTimer,
        [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString", SqlQuery = query)] IEnumerable<object> inputDocument,
        
        ILogger log)
        {

           
            log.LogInformation($"SendEmailTimer executed at: {DateTime.Now}");

            var msg = new SendGridMessage()
            {
                From = new EmailAddress("irina.burdilo@sequenceshift.com", "IB"),
                Subject = "Sending emails with Twilio SendGrid is Fun",
                PlainTextContent = "and easy to do anywhere, especially with C#",
                HtmlContent = "and easy to do anywhere, <strong>especially with C#</strong>"
            };

            msg.AddTo(new EmailAddress("ira.burdilo@gmail.com", "Iriss"));


            SendGrid.Helpers.Mail.Attachment att = new SendGrid.Helpers.Mail.Attachment
            {


                Content = Convert.ToBase64String(excelAttachment(inputDocument)),
                Filename = "Report-from.xlsx",
                Type = "application/vnd.openxmlformats-officedocument.spreadsheet.excel",
                Disposition = "attachment"

            };
            msg.AddAttachment(att);

            return msg;
        }
    }
}
