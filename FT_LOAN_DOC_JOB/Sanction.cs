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
using static SanctionApplication.Disbursal;
using Path = System.IO.Path;

namespace SanctionApplication
{
    public class Sanction
   {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        #region Sanction Logic -------------------------*********
        public static void SanctionProcedure(string connectionString, InformationModel information)
        {
            try
            {
                if (!string.IsNullOrEmpty(connectionString))
                {
                    string Agreementpath = string.Empty; string agreement_html_content = string.Empty; string disbursal_html_Content = string.Empty; string sanction_html_content = string.Empty;
                    string disbursapath = string.Empty;
                    string sanctionpath = string.Empty;
                    List<SanctionModel> sanctionLetterRs = GetSanctionLetter(information, connectionString, ConfigurationManager.AppSettings["SanctionMethodName"]);
                    foreach (var sanctionLetter in sanctionLetterRs)
                    {
                        (string pdfFilePath, string sanctionhtml_content) = GeneratePdfForSanctionLetter(sanctionLetter, sanction_html_content);
                        AllEmailBody.DispatchEmail(agreement_html_content, disbursal_html_Content, sanctionhtml_content, Agreementpath, disbursapath, pdfFilePath, sanctionLetter.email_id, sanctionLetter.user_id, sanctionLetter.lead_id, information.loan_id, sanctionLetter.name, ConfigurationManager.AppSettings["SanctionMethodName"]);
                    }
                }
                else
                {
                    logger.Error($"Please check connection string! {connectionString}");
                }
            }
            catch (Exception ex)
            {
                logger.Error("An error occurred: " + ex.Message);
            }
        }

        public static (string pdfFilePath, string sanction_html_content) GeneratePdfForSanctionLetter(SanctionModel sanctionLetter, string sanction_html_content)
        {
            string rootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
            try
            {
                //string rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                //string sanctionDocumentPath = Path.Combine(rootPath, "SanctionDocument");
                string sanctionDocumentPath = "D:\\Junoon Capital\\Fynto Job\\FINTOO_LOAN_DOC_JOB\\FT_LOAN_DOC_JOB\\SanctionDocument";

                string txtFilePath = Path.Combine(sanctionDocumentPath, $"FT_Sanction_letter.txt");

                logger.Info($"SanctionDocument Letter Root Path - {rootPath}");
                logger.Info($"SanctionDocument SanctionDocument Root Path - {sanctionDocumentPath}");
                logger.Info($"FT_Sanction_letter SanctionDocument Root Path - {txtFilePath}");

                if (!File.Exists(txtFilePath))
                {
                    logger.Error($"Error: TXT file not found at {txtFilePath}");
                    InformationSendEmail(sanctionLetter.loan_id, sanctionLetter.name, sanctionLetter.lead_id, "Sanction Procedure", $"Error: TXT file not found at {txtFilePath}");
                    throw new FileNotFoundException("Error: TXT file not found!");
                }

                string txtContent = File.ReadAllText(txtFilePath);
                Dictionary<string, string> replacements = GetReplacementData(sanctionLetter);

                foreach (var entry in replacements)
                {
                    txtContent = txtContent.Replace(entry.Key, entry.Value);
                }

                sanction_html_content = $@"{txtContent}";

                string userFileName = sanctionLetter.lead_id.Replace(" ", "_");
                string htmlFilePath = Path.Combine(sanctionDocumentPath, $"{userFileName}_FT_sanction_letter.html");
                string pdfFilePath = Path.Combine(sanctionDocumentPath, $"{userFileName}_FT_sanction_letter.pdf");

                logger.Info($"SanctionDocument SanctionDocument htmlFilePath Path - {htmlFilePath}");
                logger.Info($"FT_Sanction_letter SanctionDocument pdfFilePath Path - {pdfFilePath}");

                File.WriteAllText(htmlFilePath, sanction_html_content);

                using (PdfWriter writer = new PdfWriter(pdfFilePath))
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                {
                    pdfDoc.SetDefaultPageSize(PageSize.A4);
                    Document document = new Document(pdfDoc);
                    document.SetMargins(40, 40, 40, 40);
                    ConverterProperties props = new ConverterProperties();
                    HtmlConverter.ConvertToPdf(new FileStream(htmlFilePath, FileMode.Open), pdfDoc, props);
                    document.Close();
                }
                return (pdfFilePath , sanction_html_content);
            }
            catch (FileNotFoundException fnfEx)
            {
                logger.Error($"Error: File not found error:  {fnfEx.Message}");
                InformationSendEmail(sanctionLetter.loan_id, sanctionLetter.name, sanctionLetter.lead_id, "Sanction Procedure", $"Error: File not found error:  {fnfEx.Message}");
                return (string.Empty , string.Empty);
            }
            catch (Exception ex)
            {
                logger.Error($"An unexpected error occurred: {ex.Message}");
                InformationSendEmail(sanctionLetter.loan_id, sanctionLetter.name, sanctionLetter.lead_id, "Sanction Procedure", $"An unexpected error occurred: {ex.Message}");
                return (string.Empty, string.Empty);
            }
        }

        static Dictionary<string, string> GetReplacementData(SanctionModel sanctionLetter)
        {
            return new Dictionary<string, string>
            {
                { "[name]", string.IsNullOrEmpty(sanctionLetter.name) ? "N/A" : sanctionLetter.name },
                { "[loan_id]", string.IsNullOrEmpty(sanctionLetter.loan_id) ? "N/A" : sanctionLetter.loan_id },
                { "[current_date]", string.IsNullOrEmpty(sanctionLetter.current_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.current_date },
                { "[upcoming_lead_date]", string.IsNullOrEmpty(sanctionLetter.created_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.created_date },
                { "[amount]", string.IsNullOrEmpty(sanctionLetter.amount) ? "0" : sanctionLetter.amount },
                { "[tenure]", string.IsNullOrEmpty(sanctionLetter.tenure) ? "N/A" : sanctionLetter.tenure },
                { "[repayment_frequency]", string.IsNullOrEmpty(sanctionLetter.repayment_frequency) ? "N/A" : sanctionLetter.repayment_frequency },
                { "[interest_rate]", string.IsNullOrEmpty(sanctionLetter.interest_rate) ? "N/A" : sanctionLetter.interest_rate },
                { "[processing_fee]", string.IsNullOrEmpty(sanctionLetter.processing_fee) ? "0" : sanctionLetter.processing_fee },
                { "[insurance]", string.IsNullOrEmpty(sanctionLetter.insurance) ? "0" : sanctionLetter.insurance },
                { "[gst]", string.IsNullOrEmpty(sanctionLetter.gst) ? "0" : sanctionLetter.gst },
                { "[apr]", string.IsNullOrEmpty(sanctionLetter.APR) ? "0.00" : sanctionLetter.APR },
                { "[disbursed_amount]", string.IsNullOrEmpty(sanctionLetter.disbursed_amount) ? "0" : sanctionLetter.disbursed_amount },
                { "[number_of_installement]", string.IsNullOrEmpty(sanctionLetter.number_of_installement) ? "0" : sanctionLetter.number_of_installement },
                { "[installement_amount]", string.IsNullOrEmpty(sanctionLetter.installement_amount) ? "0" : sanctionLetter.installement_amount },
                { "[_tenure]", string.IsNullOrEmpty(sanctionLetter._tenure) ? "N/A" : sanctionLetter._tenure },
                { "[t_collected_amount]", string.IsNullOrEmpty(sanctionLetter.t_collected_amount) ? "0" : sanctionLetter.t_collected_amount },
                { "[total_i_collect]", string.IsNullOrEmpty(sanctionLetter.total_i_collect) ? "0" : sanctionLetter.total_i_collect },
                { "[emi_schedule_table]", string.IsNullOrEmpty(sanctionLetter.html_sanction_letter) ? "N/A" : sanctionLetter.html_sanction_letter },
                { "[address]", string.IsNullOrEmpty(sanctionLetter._address) ? "N/A" : sanctionLetter._address },
                { "[mobile_number]", string.IsNullOrEmpty(sanctionLetter.mobile_number) ? "N/A" : sanctionLetter.mobile_number },
                { "[email_id]", string.IsNullOrEmpty(sanctionLetter.email_id) ? "N/A" : sanctionLetter.email_id },
                { "[ip]", string.IsNullOrEmpty(sanctionLetter.ip) ? "N/A" : sanctionLetter.ip },
                { "[consent_otp_verified_time]", string.IsNullOrEmpty(sanctionLetter.consent_otp_verified_time) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.consent_otp_verified_time },
                { "[repayment_date]", string.IsNullOrEmpty(sanctionLetter.repayment_date) ? DateTime.Now.ToString("yyyy-MM-dd") : sanctionLetter.repayment_date },
                 { "[penal_charges]", string.IsNullOrEmpty(sanctionLetter.penal_charges) ? "N/A" : sanctionLetter.penal_charges },
                { "\n", string.Empty },
                { "\t", string.Empty },
                { "\"", "'" }
            };
        }

        public static List<SanctionModel> GetSanctionLetter(InformationModel information , string dbconnection , string method_name)
        {
            List<SanctionModel> sanctionLetters = new List<SanctionModel>();
            try
            {
                DataSet Objds = FetchSanctionLetterData(information, dbconnection);
                if (Objds?.Tables[0] != null && Objds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in Objds.Tables[0].Rows)
                    {
                        SanctionModel sanctionLetter = ProcessSanctionLetterRow(row);
                        sanctionLetters.Add(sanctionLetter);
                    }
                }
                if (Objds?.Tables[1] != null && Objds.Tables[1].Rows.Count > 0)
                {
                    foreach (SanctionModel letter in sanctionLetters)
                    {
                        string currentLeadId = letter.lead_id;
                        var filteredRows = Objds.Tables[1].AsEnumerable().Where(row => row["lead_id"].ToString() == currentLeadId).CopyToDataTable();
                        string htmlTable = GenerateRepaymentScheduleHtml(filteredRows);
                        letter.html_sanction_letter = htmlTable;
                    }
                }
            }
            catch (Exception ex)
            {
                string documentname = "Sanction";
                InformationSendEmail(information.loan_id, information.full_name, information.lead_id, documentname, ex.Message);
            }
            return sanctionLetters;
        }
        private static DataSet FetchSanctionLetterData(InformationModel information, string dbconnection)
        {
            using (var connection = new SqlConnection(dbconnection))
            {
                try
                {
                    SqlParameter[] param = new SqlParameter[1];
                    param[0] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
                    param[0].Value = information.loan_id;

                    return SqlHelper.ExecuteDataset(connection, CommandType.StoredProcedure, "USP_Sanction_letter_EXE", param);
                }
                catch (SqlException ex)
                {
                    logger.Error($"Database error occurred while fetching sanction letter details: {ex.Message}");

                    string documentname = "Sanction";
                    InformationSendEmail(
                        information.loan_id,
                        information.full_name,
                        information.lead_id,
                        documentname,
                        "Database error occurred while fetching sanction letter details."
                    );

                    // Return an empty DataSet to maintain method contract
                    return new DataSet();
                }
            }
        }
        private static SanctionModel ProcessSanctionLetterRow(DataRow row)
        {
           double LoanAmount = Convert.ToDouble(row["Loan_amount"] ?? 0);
           double cgst = Convert.ToDouble(row["cgst"] ?? 0);
           double sgst = Convert.ToDouble(row["sgst"] ?? 0);
           double igst = Convert.ToDouble(row["igst"] ?? 0);
            //string penal_charges = (row["penal_charge"] ?? 0).ToString();
            double total_gst_count = cgst + sgst + igst;
           double processing_fee = Convert.ToDouble(row["processing_fee"] ?? 0);
           double insurance = Convert.ToDouble(row["insurance_fee"] ?? 0);
           string repaymentFrequency = row["repayment_frequency"]?.ToString() ?? "Monthly";
           double disbursedAmount = 0;
           int PaymentPeriods = 1;
           double InterestRate = Convert.ToDouble(row["interest_rate"] ?? 0);
           double Payment = 0;
           int val = 0;
           double _apr = 0.00; 
           string _tenure = "0.00";
           decimal total_tenure;
          // MidpointRounding mode = MidpointRounding.ToEven;
           string tenure = row["tenure"]?.ToString();

           if (repaymentFrequency == "0")
           {
               tenure = row["tenure"]?.ToString()?.Replace("DAYS", "") ?? "0";
               _tenure = (tenure.Replace(tenure, "DAYS"));
               tenure = (tenure.Replace("DAYS", ""));
               total_tenure = Convert.ToInt32(tenure);
           }
           else
           {
               tenure = row["tenure"]?.ToString()?.Replace("Months", "") ?? "0";
               _tenure = (tenure.Replace(tenure, "Months"));
               tenure = (tenure.Replace("Months", ""));
               total_tenure = Convert.ToInt32(tenure);
           }
           _apr = 365 * InterestRate + processing_fee;
           processing_fee = LoanAmount * processing_fee / 100;
           insurance = insurance * LoanAmount / 100;
           total_gst_count = (processing_fee) * total_gst_count / 100;
           disbursedAmount = LoanAmount - (processing_fee + total_gst_count + insurance);
           if (repaymentFrequency == "30")
           {
               InterestRate = InterestRate / (12 * 100);
               repaymentFrequency = "Monthly";
               PaymentPeriods = Convert.ToInt32(total_tenure);
           }
           else if (repaymentFrequency == "7")
           {
               InterestRate = InterestRate / (52 * 100);
               total_tenure = total_tenure * 30 / 7;
               repaymentFrequency = "Weekly";
               PaymentPeriods = Convert.ToInt32(total_tenure);
           }
           else if (repaymentFrequency == "14")
           {
               InterestRate = InterestRate / (26 * 100);
               total_tenure = total_tenure * 30 / 14;
               repaymentFrequency = "Bi-Weekly";
               PaymentPeriods = Convert.ToInt32(total_tenure);
           }
           else if (repaymentFrequency == "15")
           {
               InterestRate = InterestRate / (24 * 100);
               total_tenure = total_tenure * 30 / 15;
               repaymentFrequency = "FourthNightly";
               PaymentPeriods = Convert.ToInt32(total_tenure);
           }
           else if (repaymentFrequency == "0")
           {
               InterestRate = InterestRate / 100;
               //total_tenure = total_tenure;
               total_tenure = Math.Round(total_tenure, val, MidpointRounding.ToEven);
               repaymentFrequency = "Bulletpayment";
               PaymentPeriods = Convert.ToInt32(1);
           }
           if (repaymentFrequency == "Bulletpayment")
           {
               Payment = LoanAmount * (Convert.ToDouble(InterestRate) * Convert.ToDouble(total_tenure));
               Payment = LoanAmount + Payment;
               Payment = Math.Round(Payment);
           }
           else
           {
               double emicalculation = Math.Pow(1 + InterestRate, PaymentPeriods);
               Payment = LoanAmount * InterestRate * emicalculation / (emicalculation - 1);
               Payment = Math.Round(Payment, 2);
           }

           return new SanctionModel
           {
               user_id = row["user_id"]?.ToString() ?? "N/A",
               lead_id = row["lead_id"]?.ToString() ?? "N/A",
               name = row["full_name"]?.ToString() ?? "N/A",
               current_date = row["disbursement_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
               created_date = row["created_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
               amount = LoanAmount.ToString(),
               tenure = tenure,
               repayment_frequency = repaymentFrequency,
               interest_rate = Convert.ToString(row?["interest_rate"] ?? 0),
               processing_fee = processing_fee.ToString(),
               insurance = insurance.ToString(),
               gst = total_gst_count.ToString(),
               disbursed_amount = disbursedAmount.ToString(),
               number_of_installement = PaymentPeriods.ToString(),
               installement_amount = Payment.ToString(),
               t_collected_amount = row["t_collected_amount"]?.ToString() ?? "0",
               total_i_collect = row["total_i_collect"]?.ToString() ?? "0",
               mobile_number = row["mobile_number"]?.ToString() ?? "NA",
               _address = row["address"]?.ToString() ?? "NA",
               email_id = row["email_id"]?.ToString() ?? "NA",
               _tenure = _tenure.ToString() ?? "NA",
               ip = row["IP"]?.ToString() ?? "NA",
               consent_otp_verified_time = row["consent_otp_verified_time"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
               repayment_date = row["repayment_date"]?.ToString() ?? DateTime.Now.ToString("yyyy-MM-dd"),
               loan_id = row["loan_id"]?.ToString() ?? "NA",
               APR = Convert.ToString(_apr) ?? "NA",
               penal_charges = row["penal_charge"]?.ToString() ?? "NA",
           };
       
        }

        private static string GenerateRepaymentScheduleHtml(DataTable table)
        {
            string html = @"
            <table style='width: 100%; border-collapse: collapse; margin-top: 20px;'>
                <tr>
                    <th colspan='5' style='padding: 12px; text-align: center; border: 1px solid #ddd; background-color: #f2f2f2;'>Detailed Repayment Schedule (Illustrative)</th>
                </tr>
                <tr>
                    <th style='padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;'>Installment<br/> No.</th>
                    <th style='padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;'>Outstanding Principal<br/> (in Rupees)</th>
                    <th style='padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;'>Principal<br/> (in Rupees)</th>
                    <th style='padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;'>Interest<br/> (in Rupees)</th>
                    <th style='padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;'>Installment<br/> (in Rupees)</th>
                </tr>";

            foreach (DataRow row in table.Rows)
            {
                html += $@"
                <tr>
                    <td style='padding:12px;text-align:left;border:1px solid #ddd;'>1</td>
                    <td style='padding:12px;text-align:left;border:1px solid #ddd;'>{row["remaining_amount"]}</td>
                    <td style='padding:12px;text-align:left;border:1px solid #ddd;'>{row["principl_due"]}</td>
                    <td style='padding:12px;text-align:left;border:1px solid #ddd;'>{row["interest_due"]}</td>
                    <td style='padding:12px;text-align:left;border:1px solid #ddd;'>{row["installment"]}</td>
            </tr>";
            }
            html += "</table>";
            return html.Replace("\r", "").Replace("\n", "").Replace("\"", "'");
        }
        public static void updatesanction(string user_id, string lead_id , string loan_id ,string sanction_letter)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            DataSet Objds = null;
            DataTable Objtable = new DataTable();
            SqlParameter[] param = new SqlParameter[4];

            param[0] = new SqlParameter("user_id", SqlDbType.VarChar, 10);
            param[0].Value = user_id;

            param[1] = new SqlParameter("lead_id", SqlDbType.VarChar, 10);
            param[1].Value = lead_id;

            param[2] = new SqlParameter("loan_id", SqlDbType.VarChar, 30);
            param[2].Value = loan_id;

            param[3] = new SqlParameter("sanction_letter", SqlDbType.Text);
            param[3].Value = sanction_letter;

            using (var connection = new SqlConnection(connectionString))
            {
                Objds = new DataSet();
                try
                {
                    Objds = SqlHelper.ExecuteDataset(connectionString, CommandType.StoredProcedure, "USP_update_sanction_letter_exe", param);
                }
                catch (Exception ex)
                {
                    logger.Error($"Error while update senction letter status: {ex.Message}");
                }
            }
        }
        #endregion
    }
}
