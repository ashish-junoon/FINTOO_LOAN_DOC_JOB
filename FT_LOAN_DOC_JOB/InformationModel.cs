using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SanctionApplication
{
   public class InformationModel
   {
        public string loan_id { get; set; } = string.Empty;
        public string full_name { get; set; } = string.Empty;
        public string lead_id { get; set; } = string.Empty;
        public bool disbursal_consent_sent_over_email { get; set; }
        public bool aggrement_consent_sent_over_email { get; set; }
        public bool sanction_consent_sent_over_email { get; set; }
    }
}
