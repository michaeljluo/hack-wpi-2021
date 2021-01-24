using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using System.Windows.Forms;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using System.IO;


// Install-Package MailKit -Version 2.2.0
// Install-Package MimeKit -Version 2.2.0
namespace Goat_Therapy_WPI_Hackathon
{
    class Email
    {
        
         public readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
       
        public void SendEmail(string email, string club)
        {
            Cursor.Current = Cursors.WaitCursor;
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("ProGoat Automated Email", "zarivernider@gmail.com"));
            message.To.Add(new MailboxAddress("Recipient", "zarivernider@wpi.edu"));
            message.Subject = "Automated Message to join " + club;

            var emailbod = new BodyBuilder();

            emailbod.HtmlBody = "<body style=background-color:ivory;>" + "<font face=georgia size=4>" +
@"Hello, <br>
<br>
This is an automated message saying that the sender would like to join " + club + " which can be found at the following email address: " + email + @"<br> <br>
Thank you for using the ProGoat automated software. <br>";
            message.Body = emailbod.ToMessageBody();
            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.Connect("smtp.gmail.com", 587);
                client.Authenticate("zarivernider@gmail.com", "GenericPassword1");
                


                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                // Note: only needed if the SMTP server requires authentication
                //client.Authenticate("zriverni@ctscorp.com", "password"); //CTSTestEmail

                client.Send(message);
                client.Disconnect(true);
            }
            Cursor.Current = Cursors.Default;
        }
       
       

    }
}