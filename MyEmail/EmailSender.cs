using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
namespace MyEmail
{
    public class EmailSender
    {
        private readonly string smtpServer;
        private readonly int smtpPort;
        private readonly string smtpUser;
        private readonly string smtpPass;

        public EmailSender(string smtpServer, int smtpPort, string smtpUser, string smtpPass)
        {
            this.smtpServer = smtpServer;
            this.smtpPort = smtpPort;
            this.smtpUser = smtpUser;
            this.smtpPass = smtpPass;
        }

        public void SendEmail(List<string> recipients, List<string> cc, string subject, string body, bool EnableSsl, string attachmentPath = null)
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(smtpUser);
                foreach (string recipient in recipients)
                {
                    mail.To.Add(recipient);
                }
                foreach (string recipient in cc)
                {
                    mail.CC.Add(recipient);
                }
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = false;

                // 附加檔案
                if (!string.IsNullOrEmpty(attachmentPath))
                {
                    Attachment attachment = new Attachment(attachmentPath);
                    mail.Attachments.Add(attachment);
                }

                using (SmtpClient smtp = new SmtpClient(smtpServer, smtpPort))
                {
                    smtp.Credentials = new NetworkCredential(smtpUser, smtpPass);
                    smtp.EnableSsl = EnableSsl; // 如果你的SMTP服务器使用SSL则设置为true
                    smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtp.Send(mail);
                }
            }
        }
    }
}
