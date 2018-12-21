using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;



namespace Tools
{
    public class Mail
    {
        
        public static bool SendMailSMTP(string _to, string _subject, string _body, bool isBodyHTML = false, List<string> _attachmentsPath = null, string _cc = "")
        {
            try
            {

                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("application.blabla.net");
                MailMessage mail = new MailMessage();

                mail.From = new MailAddress("blablabla@gmail.com", "BLABLA");
                mail.To.Add(_to);
                mail.Subject = "[BLABLA]" + _subject;
                mail.Body = _body;
                mail.IsBodyHtml = isBodyHTML;
                if (_attachmentsPath != null) { _attachmentsPath.ForEach(x => mail.Attachments.Add(new Attachment(x))); }
                if (!string.IsNullOrEmpty(_cc)) { mail.CC.Add(_cc); }

                client.Send(mail);

            }catch(Exception e)
            {
                Logger.Instance.Log("Unable to send email through SMTP: " + e.Message);
                return false;
            }

            return true;            
            
        }


    }
}
