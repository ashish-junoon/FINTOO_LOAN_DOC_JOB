using System;
using System.IO;
using System.Collections.Generic;
using iText.Html2pdf;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Layout;
using Path = System.IO.Path;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using LMS_DL;
using static SanctionApplication.Disbursal;
using System.Data.Common;
using System.Globalization;

namespace SanctionApplication
{
    public class Agreement
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        #region Agreement Logic -------------------------*****************
        public static void AgreementProcedure(string connectionString, InformationModel information)
        {
            string documentname = "Agreement document";
            string rootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
            try
            {
                if (connectionString != null || connectionString != "")
                {
                    string Agreementpath = string.Empty; string agreement_html_content = string.Empty; string disbursal_html_Content = string.Empty; string sanction_html_content = string.Empty;
                    string disbursapath = string.Empty;
                    string sanctionpath = string.Empty;
                   
                    //string rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                    string agreementDocumentPath = Path.Combine(rootPath, "AgreementDocument");
                    logger.Info($"rootPath path - {rootPath}");
                    logger.Info($"agreementDocumentPath path - {agreementDocumentPath}");
                    List<AgreementModel> agreementLetters = AgreementGetAgreementLetters(information, connectionString, ConfigurationManager.AppSettings["AgreementMethodName"]);
                    foreach (var _agreement in agreementLetters)
                    {
                        string txtFilePath = Path.Combine(agreementDocumentPath, $"{_agreement.product_name}_Agreement_letter.txt");
                        logger.Info($"txtFilePath path - {txtFilePath}");
                        if (File.Exists(txtFilePath))
                        {
                            string txtContent = File.ReadAllText(txtFilePath);
                            Dictionary<string, string> replacements = AgreementGetReplacementData(_agreement);
                            foreach (var entry in replacements)
                            {
                                txtContent = txtContent.Replace(entry.Key, entry.Value);
                            }
                            agreement_html_content = $@"{txtContent}";
                            string userFileName = _agreement.lead_id.Replace(" ", "_");
                            string product_name = _agreement.product_name.Replace(" ", "_");
                            string htmlFilePath = Path.Combine(agreementDocumentPath, $"{userFileName}_{product_name}_Agreement_letter.html");
                            File.WriteAllText(htmlFilePath, agreement_html_content);
                            Agreementpath = Path.Combine(agreementDocumentPath, $"{userFileName}_{product_name}_Agreement_letter.pdf");
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
                            AllEmailBody.DispatchEmail(agreement_html_content , disbursal_html_Content, sanction_html_content , Agreementpath, disbursapath, sanctionpath, _agreement.email_id, _agreement.user_id, _agreement.lead_id, information.loan_id, _agreement.name, ConfigurationManager.AppSettings["AgreementMethodName"]);
                        }
                        else
                        {
                            logger.Error($"Error: TXT file path not found!{txtFilePath}");
                            InformationSendEmail(information.loan_id, information.full_name, information.lead_id, documentname, $"Error: TXT file path not found!{txtFilePath}");
                        }
                    }
                }
                else
                {
                    logger.Error($"Please check connection string !! {connectionString}");
                    InformationSendEmail(information.loan_id, information.full_name, information.lead_id, documentname, $"Please check connection string !! {connectionString}");
                }
            }
            catch (Exception ex)
            {
                documentname = "Agreement document";
                logger.Error($"Error while executing AgreementProcedure method !!{ex.Message}");
                InformationSendEmail(information.loan_id, information.full_name, information.lead_id, documentname, ex.Message);
            }
        }
        public static Dictionary<string, string> AgreementGetReplacementData(AgreementModel sanctionLetter)
        {
            return new Dictionary<string, string>
            {
                { "[name]", string.IsNullOrEmpty(sanctionLetter.name) ? "N/A" : sanctionLetter.name },
                //{ "[father_name]", string.IsNullOrEmpty(sanctionLetter.father_name) ? "N/A" : sanctionLetter.father_name },
                { "[address]", string.IsNullOrEmpty(sanctionLetter.address) ? "N/A" : sanctionLetter.address },
                { "[applicant_company_name]", string.IsNullOrEmpty(sanctionLetter.company_name) ? "N/A" : sanctionLetter.company_name },
                { "[company_address]", string.IsNullOrEmpty(sanctionLetter.company_address) ? "N/A" : sanctionLetter.company_address },
                { "[executionday]", sanctionLetter.executionday > 0 ? sanctionLetter.executionday.ToString() : "N/A" },
                { "[executionmonthName]", string.IsNullOrEmpty(sanctionLetter.executionmonthName) ? "N/A" : sanctionLetter.executionmonthName },
                { "[executionyear]", sanctionLetter.executionyear > 0 ? sanctionLetter.executionyear.ToString() : "N/A" },
                { "[effectiveday]", sanctionLetter.effectiveday > 0 ? sanctionLetter.effectiveday.ToString() : "N/A" },
                { "[interest_rate]", sanctionLetter.interest_rate.ToString("F2", CultureInfo.InvariantCulture) },
                { "[effectivemonthName]", string.IsNullOrEmpty(sanctionLetter.effectivemonthName) ? "N/A" : sanctionLetter.effectivemonthName },
                { "[effectiveyear]", sanctionLetter.effectiveyear > 0 ? sanctionLetter.effectiveyear.ToString() : "N/A" },
                { "[effective_date]", string.IsNullOrEmpty(sanctionLetter.effective_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.effective_date },
                { "[execution_date]", string.IsNullOrEmpty(sanctionLetter.execution_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.execution_date },
                { "[email_id]", string.IsNullOrEmpty(sanctionLetter.email_id) ? "N/A" : sanctionLetter.email_id },
                { "[mobile_number]", string.IsNullOrEmpty(sanctionLetter.mobile_number) ? "N/A" : sanctionLetter.mobile_number },
                { "[consent_otp_verified_time]", string.IsNullOrEmpty(sanctionLetter.consent_otp_verified_time) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.consent_otp_verified_time },
                { "[loan_id]", string.IsNullOrEmpty(sanctionLetter.loan_id) ? "N/A" : sanctionLetter.loan_id },
                { "[repayment_date]", string.IsNullOrEmpty(sanctionLetter.repayment_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.repayment_date },
                { "[sanction_amount]", sanctionLetter.loan_amount.ToString("F2", CultureInfo.InvariantCulture) },
                { "[tenure]", string.IsNullOrEmpty(sanctionLetter.tenure) ? "N/A" : sanctionLetter.tenure },
            };
        }

        public static List<AgreementModel> AgreementGetAgreementLetters(InformationModel informationModel, string dbconnection, string method_name)
        {
            List<AgreementModel> agreementLetters = new List<AgreementModel>();
            DataSet objDs = null;
            try
            {
                using (var connection = new SqlConnection(dbconnection))
                {
                    SqlParameter[] param = new SqlParameter[1];

                    param[0] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
                    param[0].Value = informationModel.loan_id;

                    objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_Agreement_letter_EXE", param);
                }
                if (objDs?.Tables[0] != null && objDs.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in objDs.Tables[0].Rows)
                    {
                        AgreementModel sanctionLetter = new AgreementModel
                        {
                            name = row["full_name"]?.ToString() ?? "N/A",
                            user_id = row["user_id"]?.ToString() ?? "N/A",
                            lead_id = row["lead_id"]?.ToString() ?? "N/A",
                            company_id = row["company_id"]?.ToString() ?? "N/A",
                            product_name = row["product_name"]?.ToString() ?? "N/A",
                            address = row["address"]?.ToString() ?? "N/A",
                            company_name = row["company_name"]?.ToString() ?? "N/A",
                            interest_rate = Convert.ToDecimal(row["interest_rate"]?.ToString()),
                            execution_date = row["execution_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                            effective_date = row["effective_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                            executionday = DateTime.TryParse(row["execution_date"]?.ToString(), out DateTime executionDate) ? executionDate.Day : 0,
                            executionmonthName = DateTime.TryParse(row["execution_date"]?.ToString(), out executionDate) ? executionDate.ToString("MMMM") : "N/A",
                            executionyear = DateTime.TryParse(row["execution_date"]?.ToString(), out executionDate) ? executionDate.Year : 0,
                            effectiveday = DateTime.TryParse(row["effective_date"]?.ToString(), out DateTime effectiveDate) ? effectiveDate.Day : 0,
                            effectivemonthName = DateTime.TryParse(row["effective_date"]?.ToString(), out effectiveDate) ? effectiveDate.ToString("MMMM") : "N/A",
                            effectiveyear = DateTime.TryParse(row["effective_date"]?.ToString(), out effectiveDate) ? effectiveDate.Year : 0,
                            email_id = row["email_id"]?.ToString() ?? "N/A",
                            mobile_number = row["mobile_number"]?.ToString() ?? "NA",
                            consent_otp_verified_time = row["consent_otp_verified_time"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                            ip = row["IP"]?.ToString() ?? "N/A" ,
                            loan_id = row["loan_id"]?.ToString() ?? "N/A",
                            tenure = row["tenure"]?.ToString() ?? "N/A",
                            loan_amount = Convert.ToDecimal(row["loan_amount"]),
                            repayment_date = row["repayment_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
                        };
                        agreementLetters.Add(sanctionLetter);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error while fetching AgreementGetAgreementLetters method !!{ex.Message}");
                string documentname = "Agreement";
                InformationSendEmail(informationModel.loan_id, informationModel.full_name, informationModel.lead_id, documentname, ex.Message);

            }
            return agreementLetters;
        }
        public static void Agreementupdateagreement(string user_id, string lead_id , string loan_id , string agreement_html_content)
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

            param[3] = new SqlParameter("agreement_html_content", SqlDbType.Text);
            param[3].Value = agreement_html_content;

            using (var connection = new SqlConnection(connectionString))
            {
                Objds = new DataSet();
                try
                {
                    Objds = SqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, "USP_update_agreement_letter_exe", param);
                }
                catch (Exception ex)
                {
                    string documentname = "Agreement update agreement";
                    logger.Error($"Error while updating and executing Agreementupdateagreement method !!{ex.Message}");
                    InformationSendEmail(loan_id, "", lead_id, documentname, ex.Message);

                }
            }
        }
        #endregion
    }
}
