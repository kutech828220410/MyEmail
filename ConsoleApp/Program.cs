using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyEmail;
namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> recipients = new List<string>
            {
                "hson_evan@outlook.com",
            };

            string subject = "Your Account Password has been Changed";
            string body = "Dear User,\n\nYour account password has been successfully changed.\n\nBest regards,\nYour Company";
            string attachmentPath = ""; // 替换为附件的实际路径

            try
            {
                // 使用你的 Outlook 帐号信息初始化 EmailSender
                EmailSender emailSender = new EmailSender("smtp-mail.outlook.com", 587, "hson-service@outlook.com", "KuT1Ch@75511");
                emailSender.SendEmail(recipients, subject, body, true);
                Console.WriteLine("Emails sent successfully.");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending emails: {ex.Message}");
                Console.ReadKey();
            }
        }
    }
}
