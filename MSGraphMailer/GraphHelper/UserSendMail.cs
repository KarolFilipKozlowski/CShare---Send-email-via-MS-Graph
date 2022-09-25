using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MSGraphMailer.GraphHelper
{
    public class UserSendMail
    {
        /// <summary>
        /// Send email with MS Graph sendMail
        /// https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
        /// </summary>
        /// <param name="graphClient">GraphServiceClient</param>
        /// <param name="sernder">The account that is actually used to generate the message.</param>
        /// <param name="toRecipients">The To: recipients for the message.</param>
        /// <param name="subject">The subject of the message.</param>
        /// <param name="boodyContent">The body of the message.</param>
        /// <param name="bodyContentType">Type body of the message - text or html</param>
        /// <param name="replyTo">The email addresses to use when replying.</param>
        /// <param name="ccRecipients">The Cc: recipients for the message.</param>
        /// <param name="bccRecipients">The Bcc: recipients for the message.</param>
        /// <param name="importance">The importance of the message. The possible values are: low, normal, and high.</param>
        /// <param name="attachments">List attachments paths.</param>
        /// <param name="saveToSentItems">Save message in sent folder.</param>
        /// <returns>If true == email has been sent.</returns>
        public static bool SendMail(GraphServiceClient graphClient, string sernder,
             string[] toRecipients,
             string subject, string boodyContent, BodyType bodyContentType = BodyType.Text,
             string replyTo = null,
             string[] ccRecipients = null,
             string[] bccRecipients = null,
             Importance importance = Importance.Normal,
             string[] attachments = null,
             bool saveToSentItems = false)
        {
            bool sendMailStatus = false;
            try
            {
                var message = new Message();
                message.Subject = subject;
                message.Body = new ItemBody
                {
                    ContentType = bodyContentType,
                    Content = boodyContent
                };

                var _toRecipients = new List<Recipient>();
                foreach (var recipient in toRecipients)
                {
                    _toRecipients.Add(new Recipient
                    {
                        EmailAddress = new EmailAddress { Address = recipient }
                    });
                }
                message.ToRecipients = _toRecipients;

                if (ccRecipients != null)
                {
                    var _ccRecipients = new List<Recipient>();
                    foreach (var recipient in ccRecipients)
                    {
                        _ccRecipients.Add(new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = recipient }
                        });
                    }
                    message.CcRecipients = _ccRecipients;
                }

                if (bccRecipients != null)
                {
                    var _bccRecipients = new List<Recipient>();
                    foreach (var recipient in bccRecipients)
                    {
                        _bccRecipients.Add(new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = recipient }
                        });
                    }
                    message.BccRecipients = _bccRecipients;
                }

                if (replyTo != null)
                {
                    message.ReplyTo = new List<Recipient>() { new Recipient { EmailAddress = new EmailAddress { Address = replyTo } } };
                }
                if (attachments != null)
                {
                    var _attachments = new MessageAttachmentsCollectionPage();
                    foreach (var attachment in attachments)
                    {
                        try
                        {
                            _attachments.Add(new FileAttachment
                            {
                                ContentBytes = System.IO.File.ReadAllBytes(attachment),
                                Name = System.IO.Path.GetFileName(attachment)
                            });
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Failed add attachment: {ex.Message}");
                            Console.ResetColor();
                        }
                    }
                    message.Attachments = _attachments;
                }
                message.Importance = importance;

                var sendMail = graphClient.Users[sernder].SendMail(message, saveToSentItems).Request().PostResponseAsync().Result;
                if (sendMail.StatusCode == System.Net.HttpStatusCode.Accepted)
                {
                    sendMailStatus = true;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Email has been sent.");
                    Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
            return sendMailStatus;
        }
    }
}