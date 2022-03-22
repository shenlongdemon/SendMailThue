using EASendMail;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Google.Apis.Gmail.v1.Data;

namespace SendMailThue
{
    public class EmailUtils
    {

        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "SendMailThueApp";
        public static async Task<bool> SendGMail(string toEmail, string subject, string body)
        {

            byte[] credentialFile = Properties.Resources.credential;
            UserCredential credential;
            //read your credentials file
            using (Stream stream = new MemoryStream(credentialFile))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None).Result;
            }
            

            string message = $"To: {toEmail}\r\nSubject: {subject}\r\nContent-Type: text/html;charset=utf-8\r\n\r\n<h1>{body}</h1>";
            var data = Encoding.UTF8.GetBytes(message);
            var base64Msg = Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", ""); ;
            //call your gmail service
            var service = new GmailService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName });
            var msg = new Google.Apis.Gmail.v1.Data.Message();
            msg.Raw = base64Msg;
            service.Users.Messages.Send(msg, "me").Execute();
            return true;
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

       public static async Task SendMailAsync1(string userId, string token, string fromEmail, string toEmail, string subject, string body, List<string> attachments)
        {
            try
            {
                // Gmail API server address
                using (var client = new MailKit.Net.Smtp.SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);

                    var oauth2 = new SaslMechanismOAuth2("tranphuochungid@gmail.com", token);
                    client.Authenticate(oauth2);

                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("Joey Tribbiani", fromEmail));
                    message.To.Add(new MailboxAddress("Mrs. Chanandler Bong", toEmail));
                    message.Subject = "How you doin'?";

                    message.Body = new TextPart("plain")
                    {
                        Text = @"Hey Chandler"
                    };
                    await client.SendAsync(message);
                    client.Disconnect(true);
                }
            }
            catch (Exception ep)
            {
                Console.WriteLine("Exception: {0}", ep.Message);
            }
        }

        internal static void SendMail(string token)
        {
            
        }

    }
}

