using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using LMS_DL;
using System.Data.SqlClient;
using System.Data;
using SanctionApplication;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using System.Linq;
using NLog;

class Program
{
    private static readonly NLog.Logger logger = LogManager.GetCurrentClassLogger();

    static void Main()
    {
        string RootPath = ConfigurationManager.AppSettings["RootDirectory"].ToString();
        logger.Info($"Application started...{DateTime.Now}");  
        try
        {
           // string rootPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            string[] folders = {
               Path.Combine(RootPath, "AgreementDocument"),
               Path.Combine(RootPath, "DisbursalDocument"),
               Path.Combine(RootPath, "SanctionDocument")
           };
            DeleteFiles(folders);
            RunMainCode();
        }
        catch (Exception ex)
        {
            logger.Error($"Error while Executing code - {ex.Message}");
            //Console.WriteLine($"Unexpected error in Main method: {ex.Message}");
        }
        logger.Info($"Application end...{DateTime.Now}");
    }

    static void DeleteFiles(string[] folders)
    {
        try
        {
            foreach (var folder in folders)
            {
                if (Directory.Exists(folder))
                {
                    foreach (var file in Directory.GetFiles(folder, "*.html"))
                    {
                        File.Delete(file);
                    }
                    foreach (var file in Directory.GetFiles(folder, "*.pdf"))
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    logger.Error($"Folder does not exist - {folder}");
                }
            }
            Console.WriteLine($"Deleted");
        }
        catch (Exception ex)
        {
            logger.Error($"Error while deleting files from folder: {ex.Message}");
        }
    }

    static void RunMainCode()
    {
        try
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CrediCash_Dev"].ConnectionString;
            if (!string.IsNullOrEmpty(connectionString))
            {
                List<InformationModel> informationModels = GetIncompleteData(connectionString);
                if (informationModels.Count > 0)
                {
                    foreach (var information in informationModels)
                    {
                        if (!string.IsNullOrEmpty(information.loan_id) && !string.IsNullOrEmpty(information.lead_id))
                        {
                            bool combinedServiceEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings["CombinedService"]);
                            bool agreementServiceEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings["AgreementService"]);
                            bool sanctionServiceEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings["SanctionService"]);
                            bool disbursalServiceEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings["DisbursalService"]);
                            if (combinedServiceEnabled)
                            {
                                if (!information.sanction_consent_sent_over_email && !information.disbursal_consent_sent_over_email && !information.aggrement_consent_sent_over_email)
                                {
                                    logger.Info($"Send combined process start: {information.loan_id}");
                                    SendEmailWithCombined.SendEmailCombined(connectionString, information);
                                    logger.Info($"Send combined process end: {information.loan_id}");
                                }
                                else
                                {
                                    HandleIndividualServices(information, agreementServiceEnabled, sanctionServiceEnabled, disbursalServiceEnabled, connectionString);
                                }
                            }
                            else
                            {
                                logger.Info($"Send indivisual process start: {information.loan_id}");
                                HandleIndividualServices(information, agreementServiceEnabled, sanctionServiceEnabled, disbursalServiceEnabled, connectionString);
                                logger.Info($"Send indivisual process end: {information.loan_id}");
                            }
                        }
                    }
                }
                else
                {
                    logger.Error($"Information: NO any data found for sending email.{informationModels.Count}");
                }
            }
            else
            {
                logger.Error($"Connection string is empty or missing..{connectionString}");
            }
        }
        catch (Exception ex)
        {
            logger.Error($"Error while running main code logic: {ex.Message}");
        }
    }

    public static void HandleIndividualServices(InformationModel information, bool agreementServiceEnabled, bool sanctionServiceEnabled, bool disbursalServiceEnabled, string connectionString)
    {
        try
        {
            if (!information.aggrement_consent_sent_over_email && agreementServiceEnabled)
            {
                logger.Info($"Agreement process start: {information.loan_id}");
                Agreement.AgreementProcedure(connectionString, information);
                logger.Info($"Agreement process end: {information.loan_id}");
            }
            if (!information.sanction_consent_sent_over_email && sanctionServiceEnabled)
            {
                logger.Info($"Sanction process start: {information.loan_id}");
                Sanction.SanctionProcedure(connectionString, information);
                logger.Info($"Sanction process end: {information.loan_id}");
            }
            if (!information.disbursal_consent_sent_over_email && disbursalServiceEnabled)
            {
                logger.Info($"Disbursal process start: {information.loan_id}");
                Disbursal.DisbursalProcedure(connectionString, information);
                logger.Info($"Disbursal process end: {information.loan_id}");
            }
        }
        catch (Exception ex)
        {
            logger.Error($"Error in HandleIndividualServices: {ex.Message}");
            Disbursal.InformationSendEmail(information.loan_id, information.full_name, information.lead_id, "Handle Individual Services", $"Error in HandleIndividualServices: {ex.Message}");
        }
    }

    public static List<InformationModel> GetIncompleteData(string dbconnection)
    {
        List<InformationModel> informationModels = new List<InformationModel>();
        DataSet objDs = null;

        try
        {
            using (var connection = new SqlConnection(dbconnection))
            {
                SqlParameter[] param = new SqlParameter[1];
                objDs = SqlHelper.ExecuteDataset(dbconnection, CommandType.StoredProcedure, "USP_GetIncompleteData", param);
            }
            {
                foreach (DataRow row in objDs.Tables[0].Rows)
                {
                    InformationModel information = new InformationModel
                    {
                        sanction_consent_sent_over_email = Convert.ToBoolean(row["sanction_consent_sent_over_email"]),
                        disbursal_consent_sent_over_email = Convert.ToBoolean(row["disbursal_consent_sent_over_email"]),
                        aggrement_consent_sent_over_email = Convert.ToBoolean(row["aggrement_consent_sent_over_email"]),
                        full_name = Convert.ToString(row["full_name"]),
                        loan_id = Convert.ToString(row["loan_id"]),
                        lead_id = Convert.ToString(row["lead_id"]),
                    };
                    informationModels.Add(information);
                }
            }
        }
        catch (Exception ex)
        {
            logger.Error($"Error fetching incomplete data: {ex.Message}");
        }

        return informationModels;
    }
}
