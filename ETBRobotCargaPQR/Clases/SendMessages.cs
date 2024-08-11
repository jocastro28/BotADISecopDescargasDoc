using System.Text;
using System.Net.Mail;
using System.Net;

namespace ETBRobotAsignarCasosPQR.Clases
{
    class SendMessages
    {
        //CONFIGURACION DE CORREO
        public const string SmtpPort = "587";
        public const string SmtpHost = "smtp-legacy.office365.com";
        public const string SmtpMail = "ubaldocastro12@hotmail.com";
        public const string SmtpPwd = "W@ldo281195";
        public void SendAMessage(string Asunto, string ContenidoMsj, string[] Destinatarios, string enviroment)
        {
            MailMessage mail = new MailMessage
            {
                From = new MailAddress(SmtpMail, "ADI - Robot descarga de contratos"),
                Subject = "("+enviroment + ") ADI-CONTRATOS " +Asunto,
                SubjectEncoding = Encoding.UTF8,
                Body = ContenidoMsj + "</br></br>",
                BodyEncoding = Encoding.UTF8,
                IsBodyHtml = true
            };
            foreach (string dest in Destinatarios)
            {
                mail.To.Add(dest);
            }
            SmtpClient smtp = new SmtpClient
            {  
                EnableSsl = true,
                Port = int.Parse(SmtpPort),
                Host = SmtpHost,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(SmtpMail, SmtpPwd)
            };
            smtp.Send(mail);
            smtp.Dispose();
        }
    }
}
