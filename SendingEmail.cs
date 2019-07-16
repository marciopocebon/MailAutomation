using System;
using System.Net.Mail;

namespace MailAutomation
{
    public class SendingEmail
    {
        // read from settings file
        private static string FromEmailAddress;
        private static string FromEmailPassword;
        private static string ToEmailAddress1;
        private static string ToEmailAddress2;
        private static string ToEmailAddress3;
        private static string SMTPString;
        private static string SMTPPort;
        private static string ErrorHandlerEmailAddress;
        private static string ErrorHandlerEmailAddressSMTPString;
        private static string ErrorHandlerEmailAddressSMTPPort;
        //read from XML file
        private static string ClientEmailAddress;

        public SendingEmail()
        {
        }

        public SendingEmail(string fromEmailAddress, 
                            string fromEmailPassword, 
                            string toEmailAddress1, 
                            string toEmailAddress2, 
                            string toEmailAddress3, 
                            string sMTPString,
                            string sMTPPort,
                            string errorHandlerEmailAddress,
                            string errorHandlerEmailAddressSMTPString,
                            string errorHandlerEmailAddressSMTPPort)
        {
            FromEmailAddress = fromEmailAddress ?? throw new ArgumentNullException(nameof(fromEmailAddress));
            FromEmailPassword = fromEmailPassword ?? throw new ArgumentNullException(nameof(fromEmailPassword));
            ToEmailAddress1 = toEmailAddress1 ?? throw new ArgumentNullException(nameof(toEmailAddress1));
            ToEmailAddress2 = toEmailAddress2 ?? throw new ArgumentNullException(nameof(toEmailAddress2));
            ToEmailAddress3 = toEmailAddress3 ?? throw new ArgumentNullException(nameof(toEmailAddress3));
            SMTPString = sMTPString ?? throw new ArgumentNullException(nameof(sMTPString));
            SMTPPort = sMTPPort ?? throw new ArgumentNullException(nameof(sMTPPort));
            ErrorHandlerEmailAddress = errorHandlerEmailAddress ?? throw new ArgumentNullException(nameof(errorHandlerEmailAddress));
            ErrorHandlerEmailAddressSMTPString = errorHandlerEmailAddressSMTPString ?? throw new ArgumentNullException(nameof(errorHandlerEmailAddressSMTPString));
            ErrorHandlerEmailAddressSMTPPort = errorHandlerEmailAddressSMTPPort ?? throw new ArgumentNullException(nameof(errorHandlerEmailAddressSMTPPort));

        }

        public void SendEmail(int InterviewID, string PathToPDFWithFileName, string ClientEmailAddress)
        {
            string from, to;
            SmtpClient MailServer;
            MailMessage msg;
            string ToEmailAddress;
            SendingEmail.ClientEmailAddress = ClientEmailAddress;
            try
            {
                MailServer = new SmtpClient(SMTPString, int.Parse(SMTPPort));
                MailServer.EnableSsl = true;
                MailServer.Credentials = new System.Net.NetworkCredential(FromEmailAddress, FromEmailPassword);
            
                ToEmailAddress = ClientEmailAddress;
                //Choose the destination Email Address

                from = FromEmailAddress;
                to = ToEmailAddress;
                msg = new MailMessage(from, to);
                msg.Subject = "Report from Interview Survey Data Entry";
                msg.Body = "The PDF Report is attached.";
                PathToPDFWithFileName = PathToPDFWithFileName.Replace(".pdf", "") + InterviewID + ".pdf";
                System.Console.WriteLine(PathToPDFWithFileName);
                msg.Attachments.Add(new Attachment(PathToPDFWithFileName));
                MailServer.Send(msg);
                System.Console.WriteLine($"Message Sent Successfully to { ClientEmailAddress }.");
            }
            catch (Exception ex)
            {
                MailServer = new SmtpClient(ErrorHandlerEmailAddressSMTPString, 
                                  int.Parse(ErrorHandlerEmailAddressSMTPPort));
                MailServer.EnableSsl = true;

                from = FromEmailAddress;
                to = ErrorHandlerEmailAddress;

                MailServer.Credentials = new System.Net.NetworkCredential(from, to);


                String Error = "Unable to send email. Error : " + ex;
                Console.WriteLine(Error);

                msg = new MailMessage(from, to);

                msg.Subject = "Unable to send the message";
                msg.Body = $"The message to {ClientEmailAddress} was not sent. Error:{Error}.";
                msg.Body += $" The SMTPString is { SMTPString } and the SMTP port is { SMTPPort }. Please check all values";
                msg.Attachments.Add(new Attachment(PathToPDFWithFileName));
                MailServer.Send(msg);
                System.Console.WriteLine($"Message not sent to { ClientEmailAddress }. Error Message Sent to {to}.");



            }
        }
    }
}