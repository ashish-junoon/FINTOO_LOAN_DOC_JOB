using System.Configuration;
using System.Net.Mail;
using System.Net;
using System;
using System.IO;
using System.Collections.Generic;
using static SanctionApplication.Disbursal;
using System.Text.RegularExpressions;
using static System.Net.WebRequestMethods;

namespace SanctionApplication
{
    public class AllEmailBody
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public static string CombinedSingleEmailTemplate(string method_name , string name)
        {
            string dynamicContent = "";
            string customer_name = name;
            if (method_name == "Disbursal")
            {
                dynamicContent = "<ul>" +
                                 "   <li>Disbursal Letter – A confirmation of your credit Line disbursal.</li>" +
                                 "</ul>";
            }
            else if (method_name == "Sanction")
            {
                dynamicContent = "<ul>" +
                                 "   <li>Sanction Letter – Outlining the key terms and conditions of your credit disbursal.</li>" +
                                 "</ul>";
            }
            else if (method_name == "Agreement")
            {
                dynamicContent = "<ul>" +
                                 "   <li>Agreement Letter – Outlining the key terms and conditions of your credit line.</li>" +
                                 "</ul>";
            }
            else if (method_name == "Combined")
            {
                dynamicContent = "<ul>" +
                                 "   <li>Sanction Letter – Outlining the key terms and conditions of your credit disbursal.</li>" +
                                 "   <li>Agreement Letter – Outlining the key terms and conditions of your credit line.</li>" +
                                 "   <li>Disbursal Letter – A confirmation of your credit Line disbursal.</li>" +
                                 "</ul>";
            }
            
            else
            {
                dynamicContent = "<ul>" +
                                 "   <li>Agreement Letter – Outlining the key terms and conditions of your credit line.</li>" +
                                 "</ul>";
            }
            return (
                " <!DOCTYPE html>" +
                " <html>" +
                "   <body style='font-family: Arial, sans-serif; line-height: 1.5; color: #333; margin: 0; padding: 0;'>" +
                "       <div style='margin: 20px auto; padding: 20px;'>" +
                "           <div style='font-size: 18px; font-weight: bold; margin-bottom: 20px;'>Dear "+ customer_name + ",</div>" +
                "               <p>Greetings from "+ ConfigurationManager.AppSettings["Product"] + "!</p>" +
                "               <p>We are pleased to inform you that your Credit Line against Salary has been successfully disbursed.</p>" +
                "               <p>Please find the attached documents for your reference:</p>" +
                                dynamicContent +
                "               <p>We sincerely thank you for placing your trust in " + ConfigurationManager.AppSettings["product_name"] + " Your relationship manager remains available for any support or clarification you may require.</p>" +
               // "               <p><strong>Relationship Manager: " + ConfigurationManager.AppSettings["RelationshipManager"] + "</strong><br>" +
                "               <strong>Mobile No.:</strong> " + ConfigurationManager.AppSettings["MobileNo"] + " <br>" +
                "               <strong>E-Mail ID:</strong> " + ConfigurationManager.AppSettings["EMail_ID"] + "</p>" +
                "               <p>We value your association with us and look forward to serving you again.</p>" +
                "               <p>Best Regards,<br>" +
                "               <strong>Team "+ ConfigurationManager.AppSettings["Product"] + "</strong><br>" +
                "               <strong>(Lending Partner : " + ConfigurationManager.AppSettings["LendingPartner"] + ")</strong></p>" +
                "               <p style='font-size: 14px; color: #666; border-top: 1px solid #ddd; padding-top: 10px;'>" +
                "               <em>This is an auto-generated email. Please do not reply to this address.</em>" +
                "           </p>" +
                "       </div>" +
                "   </body>" +
                " </html>"
            );
        }
        public static List<string> SendEmail(List<string> pdfAttachmentPaths, string recipientEmail, string user_id, string lead_id, string loan_id,string disbursalProcedure, string emailBody, string subject, string name ,  string ProductCode = "PU")
        {
            List<string> successfullyAttachedFiles = new List<string>();
            try
            {
                string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
                int smtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
                string senderEmail = ConfigurationManager.AppSettings["EmailFrom"];
                string senderPassword = ConfigurationManager.AppSettings["EmailPassword"];
                string EmailCC = ConfigurationManager.AppSettings["Emailcc"];
                string EmailBcc = ConfigurationManager.AppSettings["EmailBcc"];

                string htmlFilePath = Path.GetTempFileName() + ".html";
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(senderEmail, ConfigurationManager.AppSettings["EmailName"]),
                    Subject = subject+" Loan_ID - "+ loan_id,
                    Body = emailBody,
                    IsBodyHtml = true
                };
                if (Convert.ToBoolean(ConfigurationManager.AppSettings["IsDevelopment"]))
                {
                    mail.To.Add(new MailAddress(ConfigurationManager.AppSettings["DevelopmentEmail"], name));
                }
                else
                {
                    mail.To.Add(new MailAddress(recipientEmail, name));
                    if (!string.IsNullOrEmpty(EmailCC))
                    {
                        foreach (var email in EmailCC.Split('|'))
                        {
                            mail.CC.Add(email);
                        }
                    }
                    if (!string.IsNullOrEmpty(EmailBcc))
                    {
                        foreach (var email in EmailBcc.Split('|'))
                        {
                            mail.Bcc.Add(email);
                        }
                    }
                }

                foreach (var pdfPath in pdfAttachmentPaths)
                {
                    if (!string.IsNullOrEmpty(pdfPath) && System.IO.File.Exists(pdfPath))
                    {
                        mail.Attachments.Add(new Attachment(pdfPath));
                        successfullyAttachedFiles.Add(pdfPath);
                    }
                    else
                    {
                        logger.Info($"Attachment failed: {pdfPath}");
                        InformationSendEmail(loan_id, name, lead_id, $"Attachment failed: {pdfPath}", "Attachment failed");
                        //Console.WriteLine($"Attachment failed: {pdfPath}");
                    }
                }
                SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort)
                {
                    Credentials = new NetworkCredential(senderEmail, senderPassword),
                    EnableSsl = true
                };
                try
                {
                    smtpClient.Send(mail);
                    logger.Info("Email sent successfully!");
                    successfullyAttachedFiles.Add("success");
                }   
                catch (SmtpException smtpEx)
                {
                    logger.Error($"SMTP error: {smtpEx.Message}");
                    successfullyAttachedFiles.Add("Failed !!");
                    InformationSendEmail(loan_id, name, lead_id, $"SMTP error : {smtpEx.Message}", "Failed !!");
                }
                catch (Exception ex)
                {
                    logger.Error($"General error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                 logger.Error($"Error setting up email: {ex.Message}");
                InformationSendEmail(loan_id, name, lead_id, $"Error setting up email: {ex.Message}", "Failed !!");
            }
            return successfullyAttachedFiles;
        }
       
        public static void DispatchEmail(string agreement_html_content , string disbursal_htmlContent, string sanction_html_content , string agreementPath, string disbursalPath, string sanctionPath, string emailId, string userId, string leadId, string loan_id, string name , string disbursalProcedure)
        {
            List<string> pdfAttachmentPaths = new List<string>();
            if (!string.IsNullOrEmpty(agreementPath)) pdfAttachmentPaths.Add(agreementPath);
            if (!string.IsNullOrEmpty(disbursalPath)) pdfAttachmentPaths.Add(disbursalPath);
            if (!string.IsNullOrEmpty(sanctionPath)) pdfAttachmentPaths.Add(sanctionPath);
            string emailBody = CombinedSingleEmailTemplate(disbursalProcedure , name);
            string subject = disbursalProcedure == "Agreement" ? ConfigurationManager.AppSettings["AgreementSubject"] :
                             disbursalProcedure == "Sanction" ? ConfigurationManager.AppSettings["SanctionSubject"] :
                             disbursalProcedure == "Disbursal" ? ConfigurationManager.AppSettings["DisbursalSubject"] :
                             ConfigurationManager.AppSettings["CombinedEmailSubject"];

            // ***************************** Send Email Single pdf and multiple pdf ****************************************

            List<string> sentAttachments = SendEmail(pdfAttachmentPaths, emailId, userId, leadId, loan_id, disbursalProcedure, emailBody, subject , name);
            foreach (var attachment in sentAttachments)
            {
                if (!string.IsNullOrEmpty(attachment))
                {
                    // ***************************** Find Folder Name and update the database ****************************************

                    Match match = Regex.Match(attachment, @"\\(AgreementDocument|DisbursalDocument|SanctionDocument)\\", RegexOptions.IgnoreCase);
                    
                    if (match.Success)
                    {
                        string folderName = match.Groups[1].Value;
                        if (folderName == "AgreementDocument")
                        {
                            Agreement.Agreementupdateagreement(userId, leadId , loan_id , agreement_html_content);
                        }
                        if (folderName == "DisbursalDocument")
                        {
                            Disbursal.Disbursalupdate(userId, leadId , loan_id , disbursal_htmlContent);
                        }
                        if (folderName == "SanctionDocument")
                        {
                            Sanction.updatesanction(userId, leadId , loan_id , sanction_html_content);
                        }
                    }
                }
            }
            Console.WriteLine("Data update Successfully !!");
        }
    }
}
