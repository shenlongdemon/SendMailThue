using MimeKit;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using System.IO;
using System.Threading;
using Google.Apis.Gmail.v1.Data;

namespace SendMailThue
{
    public class EmailUtils
    {

        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "SendMailThueApp";
        public static void SendGMail(string toEmail, string subject, string body, List<string> attachmentFiles)
        {
            byte[] credentialFile = Properties.Resources.credential;
            UserCredential credential;
            //read your credentials file
            using (Stream stream = new MemoryStream(credentialFile))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None).Result;
            }

            var emailContent = new MimeMessage();
            emailContent.To.Add(new MailboxAddress("", toEmail));
            emailContent.Subject = subject;

            var builder = new BodyBuilder();
            // In order to reference selfie.jpg from the html text, we'll need to add it
            // to builder.LinkedResources and then use its Content-Id value in the img src.
            //var image = builder.LinkedResources.Add(@"C:\Users\ASUS\Downloads\logout.png");
            //image.ContentId = MimeUtils.GenerateMessageId();
            // Set the html version of the message text
            //builder.HtmlBody = string.Format(@"<p>Hey Alice,<br>
            //<center><img src=""cid:{0}""></center>", image.ContentId);
            builder.HtmlBody = body;
            if (attachmentFiles != null)
            {
                foreach (string file in attachmentFiles)
                {
                    builder.Attachments.Add(file);
                }
            }
            // We may also want to attach a calendar event for Monica's party...

            // Now we just need to set the message body and we're done
            emailContent.Body = builder.ToMessageBody();
            Message message = new Message();
            message.Raw = Encode(emailContent);
            var service = new GmailService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName });
            service.Users.Messages.Send(message, "me").Execute();
        }

        public static string Encode(MimeMessage mimeMessage)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                mimeMessage.WriteTo(ms);
                return Convert.ToBase64String(ms.GetBuffer())
                    .TrimEnd('=')
                    .Replace('+', '-')
                    .Replace('/', '_');
            }
        }

        static string Base64UrlEncode(string input)
        {
            var data = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }
        public static void SendMail(string account, string password, string fromEmail, string toEmail, string subject, string body, List<string> attachments)
        {
            MailMessage message = new MailMessage();
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient();
            System.Net.Mail.MailAddress fromAddress = new System.Net.Mail.MailAddress(fromEmail);
            message.From = fromAddress;
            message.To.Add(toEmail);
            message.Subject = subject;
            message.IsBodyHtml = true;
            message.Body = body;
            if (attachments != null)
            {
                foreach (string attachFile in attachments)
                {
                    System.Net.Mail.Attachment attachment;
                    attachment = new System.Net.Mail.Attachment(attachFile);
                    message.Attachments.Add(attachment);
                }
            }
            // We use gmail as our smtp client
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            //smtpClient.Port = 25 ; // if deploy
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = true;
            smtpClient.Credentials = new System.Net.NetworkCredential(
                account, password);
            smtpClient.Send(message);
        } 
    }
}

