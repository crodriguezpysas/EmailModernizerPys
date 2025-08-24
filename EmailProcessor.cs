using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using CsvHelper;
using CsvHelper.Configuration;
using EmailModernizerPys.Models;

// Para evitar ambigüedad con Task:
using Task = System.Threading.Tasks.Task;

namespace EmailModernizerPys.Services
{
    public class EmailProcessor
    {
        private readonly Action<string> _log;
        private const string StateFilePath = "C:\\Emails\\lastProcessedState.txt";
        private const string CsvFilePath = "C:\\Emails\\summary.csv";
        private readonly string _emailAccount = "eyr@procesosyservicios.net";
        private readonly string _senderEmail = "embargosyrequerimientosexternosbancocajasocial@fundaciongruposocial.co";
        private readonly DateTime _defaultStartDate = new DateTime(2025, 07, 26);
        private readonly TimeSpan _endTime = new TimeSpan(23, 59, 0);

        public EmailProcessor(Action<string> logCallback)
        {
            _log = logCallback;
        }

        public async Task ProcesarEmailsAsync()
        {
            DateTime lastProcessedDateTime = GetLastProcessedDateTime(_defaultStartDate);
            try
            {
                await ProcessEmails(_emailAccount, _senderEmail, lastProcessedDateTime, _endTime);
                _log("Procesamiento completado correctamente.");
            }
            catch (MsalException ex)
            {
                _log($"Error acquiring access token: {ex.Message}");
            }
            catch (ServiceResponseException ex) when (ex.ErrorCode == ServiceError.ErrorMailboxStoreUnavailable)
            {
                _log("The specified SMTP address has no mailbox associated with it.");
            }
            catch (ServiceResponseException ex)
            {
                _log($"Exchange service error: {ex.Message}");
            }
            catch (Exception ex)
            {
                _log($"Error: {ex.Message}");
            }
        }

        private async Task ProcessEmails(string emailAccount, string senderEmail, DateTime startDateTime, TimeSpan endTime)
        {
            var cca = ConfidentialClientApplicationBuilder
                .Create(ConfigurationManager.AppSettings["appId"])
                .WithClientSecret(ConfigurationManager.AppSettings["clientSecret"])
                .WithTenantId(ConfigurationManager.AppSettings["tenantId"])
                .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };
            var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();

            var ewsClient = new ExchangeService();
            ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
            ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, emailAccount);

            DateTime endDateTime = startDateTime.Date.Add(endTime);

            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startDateTime),
                new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, endDateTime),
                new SearchFilter.IsEqualTo(EmailMessageSchema.From, senderEmail));

            int pageSize = 50;
            ItemView itemView = new ItemView(pageSize);
            itemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

            var (lastProcessedFolderCounter, lastProcessedDateTime) = GetLastProcessedState();
            int folderCounter = lastProcessedFolderCounter + 1;

            FindItemsResults<Item> results;
            do
            {
                results = ewsClient.FindItems(WellKnownFolderName.Inbox, searchFilter, itemView);

                foreach (Item item in results.Items)
                {
                    PropertySet propertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Attachments, ItemSchema.Body, EmailMessageSchema.From, EmailMessageSchema.Subject, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived);
                    EmailMessage email = EmailMessage.Bind(ewsClient, item.Id, propertySet);

                    if (email.DateTimeReceived > lastProcessedDateTime)
                    {
                        string folderPath = Path.Combine("C:\\Emails", startDateTime.ToString("yyyyMMdd"), $"{folderCounter++}");
                        Directory.CreateDirectory(folderPath);

                        string headerInfo = $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>EYR PYS</strong></p>" +
                                            $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>De:</strong> {email.From.Address}</p>" +
                                            $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>Enviado el:</strong> {email.DateTimeReceived}</p>" +
                                            $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>Para:</strong> {string.Join(", ", email.ToRecipients.Select(r => r.Address))}</p>" +
                                            $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>Asunto:</strong> {email.Subject}</p>";
                        string attachmentInfo = $"<p style=\"font-family: 'Times New Roman'; font-size: 12pt;\"><strong>Datos Adjuntos:</strong> " + string.Join(", ", email.Attachments.Select(a => a.Name)) + "</p>";

                        string emailBody = email.Body.BodyType == BodyType.HTML ? email.Body.Text : $"<pre>{email.Body.Text}</pre>";
                        string fullHtml = $"<html><body>{headerInfo}{attachmentInfo}{emailBody}</body></html>";

                        string textPath = Path.Combine(folderPath, "0.html");
                        File.WriteAllText(textPath, fullHtml);

                        foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in email.Attachments)
                        {
                            if (attachment is FileAttachment fileAttachment)
                            {
                                string filePath = Path.Combine(folderPath, fileAttachment.Name);
                                fileAttachment.Load(filePath);
                            }
                        }

                        var emailSummary = new EmailSummary
                        {
                            From = email.From.Address,
                            DateTimeReceived = email.DateTimeReceived,
                            To = string.Join(", ", email.ToRecipients.Select(r => r.Address)),
                            Subject = email.Subject,
                            Attachments = string.Join(", ", email.Attachments.Select(a => a.Name))
                        };

                        AppendEmailSummaryToCsv(emailSummary, CsvFilePath);

                        SaveLastProcessedState(email.DateTimeReceived, folderCounter - 1);

                        _log($"Procesado: {email.Subject} ({email.DateTimeReceived})");
                    }
                }

                itemView.Offset += pageSize;

            } while (results.MoreAvailable);

            // Todas las rutas retornan correctamente:
            await Task.CompletedTask;
        }

        private (int, DateTime) GetLastProcessedState()
        {
            if (File.Exists(StateFilePath))
            {
                var stateData = File.ReadAllLines(StateFilePath);
                if (stateData.Length == 2 &&
                    DateTime.TryParse(stateData[0], out var lastProcessedDateTime) &&
                    int.TryParse(stateData[1], out var lastProcessedFolderCounter))
                {
                    return (lastProcessedFolderCounter, lastProcessedDateTime);
                }
            }
            return (0, DateTime.MinValue);
        }

        private DateTime GetLastProcessedDateTime(DateTime defaultStartDate)
        {
            if (File.Exists(StateFilePath))
            {
                var stateData = File.ReadAllLines(StateFilePath);
                if (stateData.Length == 2 && DateTime.TryParse(stateData[0], out var lastProcessedDateTime))
                {
                    return lastProcessedDateTime;
                }
            }
            return defaultStartDate;
        }

        private void SaveLastProcessedState(DateTime lastProcessedDateTime, int lastProcessedFolderCounter)
        {
            var stateData = $"{lastProcessedDateTime:yyyy-MM-ddTHH:mm:ss}\n{lastProcessedFolderCounter}";
            File.WriteAllText(StateFilePath, stateData);
        }

        private void AppendEmailSummaryToCsv(EmailSummary emailSummary, string csvPath)
        {
            var fileExists = File.Exists(csvPath);
            using (var writer = new StreamWriter(csvPath, append: true))
            using (var csv = new CsvWriter(writer, new CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)))
            {
                if (!fileExists)
                {
                    csv.WriteHeader<EmailSummary>();
                    csv.NextRecord();
                }
                csv.WriteRecord(emailSummary);
                csv.NextRecord();
            }
        }
    }
}