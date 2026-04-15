using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SanctionApplication
{
   public class SanctionModel
    {
        public string user_id { get; set; }
        public string lead_id { get; set; }
        public string name { get; set; }
        public string current_date { get; set; }
        public string created_date { get; set; }
        public string amount { get; set; }
        public string tenure { get; set; }
        public string repayment_frequency { get; set; }
        public string interest_rate { get; set; }
        public string processing_fee { get; set; }
        public string insurance { get; set; }
        public string gst { get; set; }
        public string disbursed_amount { get; set; }
        public string number_of_installement { get; set; }
        public string installement_amount { get; set; }
        public string html_sanction_letter { get; set; }
        public string _address { get; set; }
        public string mobile_number { get; set; }
        public string email_id { get; set; }
        public string t_collected_amount { get; set; }
        public string total_i_collect { get; set; }
        public string _tenure { get; set; }
        public string company_id { get; set; }
        public string product_name { get; set; }
        public string ip { get; set; }
        public string consent_otp_verified_time { get; set; }
        public string repayment_date { get; set; }
        public string loan_id { get; set; }
        public string APR { get; set; }


    }
}
