using iText.Html2pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using LMS_DL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
using Path = System.IO.Path;
namespace SanctionApplication
{
    public class Disbursal
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        #region Disbursal logic ----------------*********
        public static void DisbursalProcedure(string connectionString, InformationModel information)
        {
            string _rootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
            try
            {
                if (!string.IsNullOrEmpty(connectionString))
                {
                    string Agreementpath = string.Empty; string agreement_html_content = string.Empty;  string disbursal_html_Content = string.Empty; string sanction_html_content = string.Empty;
                    string disbursapath = string.Empty;
                    string sanctionpath = string.Empty;
                    //string _rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                    string _disbursalDocumentPath = Path.Combine(_rootPath, "DisbursalDocument");
                    logger.Info($"rootPath path - {_rootPath}");
                    logger.Info($"disbursalDocumentPath path - {_disbursalDocumentPath}");
                    List<Disbursal_letterModel> disbursal_Letters = ProcessDisbursalLetter(information, connectionString, ConfigurationManager.AppSettings["CombinedMethodName"]);
                    foreach (var disbursalLetter in disbursal_Letters)
                    {
                        string _disbursalFilePath = Path.Combine(_disbursalDocumentPath, $"{disbursalLetter.product_name}_Disbursal_letter.txt");
                        logger.Info($"disbursalFilePath path - {_disbursalFilePath}");
                        if (File.Exists(_disbursalFilePath))
                        {
                            string disbursaltextContent = File.ReadAllText(_disbursalFilePath);
                            Dictionary<string, string> replacements = DisbursalReplacementData(disbursalLetter);
                            foreach (var entry in replacements)
                            {
                                disbursaltextContent = disbursaltextContent.Replace(entry.Key, entry.Value);
                            }
                            disbursal_html_Content = $@"{disbursaltextContent}";
                            string userFileName = disbursalLetter.loan_id.Replace(" ", "_");
                            string product_name = disbursalLetter.product_name.Replace(" ", "_");
                            string htmlFilePath = Path.Combine(_disbursalDocumentPath, $"{userFileName}_{product_name}_Disbursal_letter.html");
                            logger.Info($"Disbursal_letter.html path - {htmlFilePath}");
                            File.WriteAllText(htmlFilePath, disbursal_html_Content);
                            disbursapath = Path.Combine(_disbursalDocumentPath, $"{userFileName}_{product_name}_Disbursal_letter.pdf");
                            logger.Info($"Disbursal_letter.pdf path - {disbursapath}");
                            using (PdfWriter writer = new PdfWriter(disbursapath))
                            using (PdfDocument pdfDoc = new PdfDocument(writer))
                            {
                                pdfDoc.SetDefaultPageSize(PageSize.A4);
                                Document document = new Document(pdfDoc);
                                document.SetMargins(40, 40, 40, 40);
                                ConverterProperties props = new ConverterProperties();
                                HtmlConverter.ConvertToPdf(new FileStream(htmlFilePath, FileMode.Open), pdfDoc, props);
                                document.Close();
                            }
                            AllEmailBody.DispatchEmail(agreement_html_content, disbursal_html_Content, sanction_html_content, Agreementpath, disbursapath, sanctionpath, disbursalLetter.email_id, disbursalLetter.user_id, disbursalLetter.lead_id, disbursalLetter.loan_id, disbursalLetter.name, ConfigurationManager.AppSettings["DisbursalMethodName"]);
                        }
                        else
                        {
                            logger.Error($"disbursalFilePath path not found - {_disbursalFilePath}");
                            InformationSendEmail(information.loan_id, information.full_name, information.lead_id, "Disbursal Procedure", $"disbursalFilePath path not found - {_disbursalFilePath}");
                        }
                    }
                    Thread.Sleep(5000);
                }
                else
                {
                    logger.Info($"Please check connection string! : {connectionString}");
                    InformationSendEmail(information.loan_id, information.full_name, information.lead_id, "Disbursal Procedure", $"Please check connection string! : {connectionString}");

                }
            }
            catch (Exception ex)
            {
                logger.Error($"An error occurred while executing Disbursal Procedure.cs- {ex.Message}");
                InformationSendEmail(information.loan_id, information.full_name, information.lead_id, "Disbursal Procedure", $"An error occurred while executing DisbursalProcedure.cs- {ex.Message}");
            
            }
        }
        public static Dictionary<string, string> DisbursalReplacementData(Disbursal_letterModel disbursal_Letter)
        {
            try
            {
                return new Dictionary<string, string>
                {
                    { "[name]", string.IsNullOrEmpty(disbursal_Letter.name) ? "N/A" : disbursal_Letter.name },
                    { "[loan_id]", string.IsNullOrEmpty(disbursal_Letter.loan_id) ? "N/A" : disbursal_Letter.loan_id },
                    { "[disbursed_amount]", string.IsNullOrEmpty(disbursal_Letter.disbursed_amount) ? "N/A" : disbursal_Letter.disbursed_amount },
                    { "[t_collected_amount]", string.IsNullOrEmpty(disbursal_Letter.t_collected_amount) ? "N/A" : disbursal_Letter.t_collected_amount },
                    { "[interest_rate]", string.IsNullOrEmpty(disbursal_Letter.interest_rate) ? "N/A" : disbursal_Letter.interest_rate },
                    { "[tenure]", string.IsNullOrEmpty(disbursal_Letter.tenure) ? "N/A" : disbursal_Letter.tenure },
                    { "[_tenure]", string.IsNullOrEmpty(disbursal_Letter._tenure) ? "N/A" : disbursal_Letter._tenure },
                    { "[email_id]", string.IsNullOrEmpty(disbursal_Letter.email_id) ? "N/A" : disbursal_Letter.email_id },
                    { "[loan_amount]", string.IsNullOrEmpty(disbursal_Letter.loan_amount) ? "N/A" : disbursal_Letter.loan_amount },
                    { "[disbursement_date]", string.IsNullOrEmpty(disbursal_Letter.current_date) ? "N/A" : disbursal_Letter.current_date },
                    { "[repayment_date]", string.IsNullOrEmpty(disbursal_Letter.repayment_date) ? "N/A" : disbursal_Letter.repayment_date }
                };
            }
            catch (Exception ex)
            {
                string documentname = "Disbursal Replacement Data";
                logger.Error($"Error in DisbursalReplacementData: {ex.Message}");
                InformationSendEmail(disbursal_Letter.loan_id, disbursal_Letter.name, disbursal_Letter.lead_id, documentname, $"Error in DisbursalReplacementData: {ex.Message}");
                return new Dictionary<string, string>();
            }
        }
        public static List<Disbursal_letterModel> ProcessDisbursalLetter(InformationModel informationModels, string dbconnection, string method_name)
        {
            List<Disbursal_letterModel> disbursal_Letters = new List<Disbursal_letterModel>();
            DataSet objDs = null;
            SqlParameter[] param = new SqlParameter[1];
            try
            {
                param[0] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
                param[0].Value = informationModels.loan_id;
                using (var connection = new SqlConnection(dbconnection))
                {
                    objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_disbursal_letter_EXE", param);
                }
                if (objDs?.Tables[0] != null && objDs.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in objDs.Tables[0].Rows)
                    {
                        string _tenure = "0.00";
                        decimal total_tenure;
                       // MidpointRounding mode = MidpointRounding.ToEven;
                        string tenure = row["tenure"]?.ToString();
                        if (row["repayment_frequency"]?.ToString() == "0")
                        {
                            tenure = tenure.Replace("DAYS", "");
                            _tenure = tenure.Replace(tenure, "DAYS").Trim();
                        }
                        else
                        {
                            tenure = tenure.Replace("Months", "");
                            _tenure = tenure.Replace(tenure, "Months").Trim();
                        }

                        total_tenure = Convert.ToInt32(tenure);

                        Disbursal_letterModel disbursal = new Disbursal_letterModel
                        {
                            user_id = row["user_id"]?.ToString() ?? "N/A",
                            lead_id = row["lead_id"]?.ToString() ?? "N/A",
                            name = row["full_name"]?.ToString() ?? "N/A",
                            current_date = row["disbursement_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                            created_date = row["created_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                            tenure = tenure,
                            interest_rate = Convert.ToString(row?["interest_rate"] ?? 0),
                            disbursed_amount = row["disbursement_amount"]?.ToString() ?? "0",
                            t_collected_amount = row["t_collected_amount"]?.ToString() ?? "0",
                            mobile_number = row["mobile_number"]?.ToString() ?? "NA",
                            email_id = row["email_id"]?.ToString() ?? "NA",
                            _tenure = _tenure.ToString() ?? "NA",
                            loan_id = row["loan_id"]?.ToString() ?? "NA",
                            product_name = row["product_name"]?.ToString() ?? "NA",
                            loan_amount = row["Loan_amount"]?.ToString() ?? "0",
                            repayment_date = row["repayment_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),

                        };
                        disbursal_Letters.Add(disbursal);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error in ProcessDisbursalLetter: {ex.Message}");
                string documentname = "Disbursal";
                InformationSendEmail(informationModels.loan_id, informationModels.full_name, informationModels.lead_id, documentname, ex.Message);

            }
            return disbursal_Letters;
        }

        public static List<InformationModel> GetLoanID(string method_name, string dbconnection)
        {
            List<InformationModel> informationModels = new List<InformationModel>();
            DataSet objDs = null;
            try
            {
                using (var connection = new SqlConnection(dbconnection))
                {
                    SqlParameter[] param = new SqlParameter[1]; // Adjust parameters as needed
                    param[0] = new SqlParameter("method_name", SqlDbType.VarChar, 30);
                    param[0].Value = method_name;

                    objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_GetLoanID", param);
                }
                if (objDs?.Tables[0] != null && objDs.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in objDs.Tables[0].Rows)
                    {
                        InformationModel information = new InformationModel
                        {
                            loan_id = row["loan_id"]?.ToString() ?? "NA",
                            full_name = row["full_name"]?.ToString() ?? "NA",
                            lead_id = row["lead_id"]?.ToString() ?? "NA"
                        };
                        informationModels.Add(information);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error fetching Disbursal Letters Method GetLoanID - : {ex.Message}");
            }
            return informationModels;
        }
        public static void Disbursalupdate(string user_id, string lead_id , string loan_id , string disbursal_htmlContent)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CrediCash_Dev"].ConnectionString;
            DataSet Objds = null;
            DataTable Objtable = new DataTable();
            SqlParameter[] param = new SqlParameter[4];

            param[0] = new SqlParameter("user_id", SqlDbType.VarChar, 10);
            param[0].Value = user_id;

            param[1] = new SqlParameter("lead_id", SqlDbType.VarChar, 10);
            param[1].Value = lead_id;

            param[2] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
            param[2].Value = loan_id;

            param[3] = new SqlParameter("disbursal_htmlContent", SqlDbType.Text);
            param[3].Value = disbursal_htmlContent;

            using (var connection = new SqlConnection(connectionString))
            {
                Objds = new DataSet();
                try
                {
                    Objds = SqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, "USP_update_disbursal_letter_exe", param);
                }
                catch (Exception ex)
                {
                    string documentname = "Disbursal Update";
                    logger.Error($"Error fetching Disbursal UPdate Method Disbursalupdate - : {ex.Message}");
                    InformationSendEmail(loan_id, "", lead_id, documentname, $"Error fetching Disbursal UPdate Method Disbursalupdate - : {ex.Message}");
                }
            }
        }
        #endregion

        #region send Email Only Development team -----------------------------------------------

        public static void InformationSendEmail(string loan_id, string applicant_name, string applicant_lead, string documentname, string message)
        {
            try
            {
                string senderEmail = ConfigurationManager.AppSettings["InformationEmailFrom"];
                string senderPassword = ConfigurationManager.AppSettings["InformationEmailPassword"];
                string recipientEmail = ConfigurationManager.AppSettings["InformationEmailTo"];
                string Emailcc = ConfigurationManager.AppSettings["InformationEmailcc"];
                string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
                int smtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
                string htmlFilePath = Path.GetTempFileName() + ".html";
                string subject = string.Empty;
                string emailBody = string.Empty;
                if(loan_id == "" && applicant_lead == "")
                {
                    subject = $"Unable to send document process";
                    emailBody = $@"
                    <!DOCTYPE html>
                    <html>
                        <head>
                           <title>Email</title>
                        </head>
                            <body style='font-family: Arial, sans-serif; line-height: 1.5; color: #333; background-color: #f9f9f9; margin: 0; padding: 0;'>
                                <div style='margin: 20px auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px;'>
                                    <div style='font-size: 18px; font-weight: bold; margin-bottom: 20px; color: #0056b3;'>Hello Team EarlyWages,</div>
                                    <div style='margin-bottom: 20px;'>
                                        We were unable to send the <strong>{documentname}</strong> letter due to an issue with the following details:
                                        <ul style='margin: 10px 0; padding: 0;'>
                                            <li style='margin-bottom: 8px;'>Error Message: <strong>{message}</strong></li>
                                        </ul>
                                        Please address this at the earliest.
                                    </div>
                                    <div style='font-size: 14px; color: #666; border-top: 1px solid #ddd; padding-top: 10px;'>
                                        Thank you,<br>Team Fynto<br><br>
                                        <em>Disclaimer: This is a system-generated email. No reply is required.</em>
                                 </div>
                                </div>
                            </body>
                    </html>";
                }
                else
                {
                    subject = $"Unable to Send {documentname} Letter Loan ID - {loan_id}";
                    emailBody = $@"
                    <!DOCTYPE html>
                    <html>
                        <head>
                           <title>Email</title>
                        </head>
                            <body style='font-family: Arial, sans-serif; line-height: 1.5; color: #333; background-color: #f9f9f9; margin: 0; padding: 0;'>
                                <div style='margin: 20px auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px;'>
                                    <div style='font-size: 18px; font-weight: bold; margin-bottom: 20px; color: #0056b3;'>Hello Team PaisaUdhar,</div>
                                    <div style='margin-bottom: 20px;'>
                                        We were unable to send the <strong>{documentname}</strong> letter due to an issue with the following details:
                                        <ul style='margin: 10px 0; padding: 0;'>
                                            <li style='margin-bottom: 8px;'>Lead ID: <strong>{applicant_lead}</strong></li>
                                            <li style='margin-bottom: 8px;'>Loan ID: <strong>{loan_id}</strong></li>
                                            <li style='margin-bottom: 8px;'>Error Message: <strong>{message}</strong></li>
                                        </ul>
                                        Please address this at the earliest.
                                    </div>
                                    <div style='font-size: 14px; color: #666; border-top: 1px solid #ddd; padding-top: 10px;'>
                                        Thank you,<br>Team Fynto<br><br>
                                        <em>Disclaimer: This is a system-generated email. No reply is required.</em>
                                 </div>
                                </div>
                            </body>
                    </html>";
                }
                File.WriteAllText(htmlFilePath, emailBody);
                string emailHtmlContent = File.ReadAllText(htmlFilePath);
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(senderEmail),
                    Subject = subject,
                    Body = emailHtmlContent,
                    IsBodyHtml = true
                };
                if (!string.IsNullOrEmpty(recipientEmail))
                {
                    mail.To.Add(recipientEmail);
                }
                if (!string.IsNullOrEmpty(Emailcc))
                {
                    mail.CC.Add(Emailcc);
                }
                using (SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort))
                {
                    smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);
                    smtpClient.EnableSsl = true;
                    smtpClient.Send(mail);
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error while sending email: {ex.Message}");
                InformationSendEmail("", "", "", documentname, $"Error while sending email: {ex.Message}");
            }
        }

        #endregion
    }
}
