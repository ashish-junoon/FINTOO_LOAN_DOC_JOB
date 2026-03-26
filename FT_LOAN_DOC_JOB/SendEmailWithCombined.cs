using iText.Html2pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;
using static SanctionApplication.Agreement;
using static SanctionApplication.Disbursal;
using static SanctionApplication.Sanction;
using Path = System.IO.Path;



namespace SanctionApplication
{
    public static class SendEmailWithCombined
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        #region SendEmailCombined logic ----------------*********
        public static void SendEmailCombined(string connectionString, InformationModel informationModels)
        {
            string _rootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
            try
            {
                if (!string.IsNullOrEmpty(connectionString))
                {
                    //List<InformationModel> _disbursal_Letters = GetLoanID(ConfigurationManager.AppSettings["CombinedMethodName"], connectionString);
                    string Agreementpath = "";
                    string disbursapath = "";
                    string sanctionpath = "";
                    string email_id = "";
                    string user_id = "";
                    string lead_id = "";
                    string loan_id = informationModels.loan_id;
                    string customer_name = "";
                    string agreement_html_content = string.Empty; string disbursal_html_Content = string.Empty; string sanction_html_content = string.Empty;

                    List<Disbursal_letterModel> disbursal_Letters = ProcessDisbursalLetter(informationModels, connectionString, ConfigurationManager.AppSettings["CombinedMethodName"]);
                    List<AgreementModel> agreementLetters = AgreementGetAgreementLetters(informationModels, connectionString, ConfigurationManager.AppSettings["CombinedMethodName"]);
                    List<SanctionModel> sanctionLetterRs = GetSanctionLetter(informationModels, connectionString, ConfigurationManager.AppSettings["CombinedMethodName"]);
                    try
                    {
                        // string _rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                       
                        string _disbursalDocumentPath = Path.Combine(_rootPath, "DisbursalDocument");

                        logger.Info($"Disbursal Letter Root Path - {_rootPath}");
                        logger.Info($"Disbursal DocumentPath Root Path - {_disbursalDocumentPath}");
                        foreach (var disbursalLetter in disbursal_Letters)
                        {
                            email_id = disbursalLetter.email_id;
                            user_id = disbursalLetter.user_id;
                            lead_id = disbursalLetter.lead_id;
                            customer_name = disbursalLetter.name;
                            loan_id = disbursalLetter.loan_id;

                            string _disbursalFilePath = Path.Combine(_disbursalDocumentPath, $"{disbursalLetter.product_name}_Disbursal_letter.txt");
                            logger.Info($"Disbursal File Path .txt - {_disbursalFilePath}");
                            if (File.Exists(_disbursalFilePath))
                            {
                                string disbursaltextContent = File.ReadAllText(_disbursalFilePath);
                                Dictionary<string, string> replacements = Disbursal.DisbursalReplacementData(disbursalLetter);
                                foreach (var entry in replacements)
                                {
                                    disbursaltextContent = disbursaltextContent.Replace(entry.Key, entry.Value);
                                }

                                disbursal_html_Content = $@"{disbursaltextContent}";

                                string userFileName = disbursalLetter.lead_id.Replace(" ", "_");
                                string product_name = disbursalLetter.product_name.Replace(" ", "_");
                                string htmlFilePath = Path.Combine(_disbursalDocumentPath, $"{userFileName}_{product_name}_Disbursal_letter.html");

                                File.WriteAllText(htmlFilePath, disbursal_html_Content);

                                disbursapath = Path.Combine(_disbursalDocumentPath, $"{userFileName}_{product_name}_Disbursal_letter.pdf");
                                logger.Info($"_Disbursal letter File Path .pdf - {_disbursalFilePath}");
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
                            }
                            else
                            {
                                logger.Info($"Error: TXT file not found! - {_disbursalFilePath}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"An error occurred: {ex.Message}");
                    }
                    try
                    {
                        //string rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                        string agreementDocumentPath = Path.Combine(_rootPath, "AgreementDocument");

                        logger.Info($"AgreementDocument Letter Root Path - {_rootPath}");
                        logger.Info($"AgreementDocument DocumentPath Root Path - {agreementDocumentPath}");

                        foreach (var _agreement in agreementLetters)
                        {
                            email_id = _agreement.email_id;
                            user_id = _agreement.user_id;
                            lead_id = _agreement.lead_id;
                            customer_name = _agreement.name;
                            //loan_id = _agreement.loan_id;
                            string txtFilePath = Path.Combine(agreementDocumentPath, $"{_agreement.product_name}_Agreement_letter.txt");
                            if (File.Exists(txtFilePath))
                            {
                                string txtContent = File.ReadAllText(txtFilePath);
                                Dictionary<string, string> replacements = Agreement.AgreementGetReplacementData(_agreement);
                                foreach (var entry in replacements)
                                {
                                    txtContent = txtContent.Replace(entry.Key, entry.Value);
                                }

                                agreement_html_content = $@"{txtContent}";
                                string userFileName = _agreement.lead_id.Replace(" ", "_");
                                string product_name = _agreement.product_name.Replace(" ", "_");
                                string htmlFilePath = Path.Combine(agreementDocumentPath, $"{userFileName}_{product_name}_Agreement_letter.html");
                                logger.Info($"Agreement_letter.html Root Path - {htmlFilePath}");
                                File.WriteAllText(htmlFilePath, agreement_html_content);

                                Agreementpath = Path.Combine(agreementDocumentPath, $"{userFileName}_{product_name}_Agreement_letter.pdf");
                                logger.Info($"Agreement_letter.pdf Root Path - {Agreementpath}");
                                using (PdfWriter writer = new PdfWriter(Agreementpath))
                                using (PdfDocument pdfDoc = new PdfDocument(writer))
                                {
                                    pdfDoc.SetDefaultPageSize(PageSize.A4);
                                    Document document = new Document(pdfDoc);
                                    document.SetMargins(40, 40, 40, 40);
                                    ConverterProperties props = new ConverterProperties();
                                    HtmlConverter.ConvertToPdf(new FileStream(htmlFilePath, FileMode.Open), pdfDoc, props);
                                    document.Close();
                                }
                            }
                            else
                            {
                                logger.Error("Error: TXT file not found!");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred: {ex.Message}");
                    }
                    foreach (var sanctionLetter in sanctionLetterRs)
                    {
                        email_id = sanctionLetter.email_id;
                        user_id = sanctionLetter.user_id;
                        lead_id = sanctionLetter.lead_id;
                        customer_name = sanctionLetter.name;
                        (sanctionpath , sanction_html_content)  = GeneratePdfForSanctionLetter(sanctionLetter , sanction_html_content);
                        //loan_id = sanctionLetter.loan_id;
                    }
                    AllEmailBody.DispatchEmail(agreement_html_content, disbursal_html_Content, sanction_html_content, Agreementpath, disbursapath, sanctionpath, email_id, user_id, lead_id, loan_id, customer_name, ConfigurationManager.AppSettings["CombinedMethodName"]);
                    Agreementpath = string.Empty; disbursapath = string.Empty; sanctionpath = string.Empty; email_id = string.Empty; user_id = string.Empty; lead_id = string.Empty;
                    Thread.Sleep(2000);
                }
                else
                {
                    logger.Error("Please check connection string!");
                }
            }
            catch (Exception ex)
            {
                logger.Error("An error occurred: " + ex.Message);
            }
        }
       
        #endregion
    }
}
