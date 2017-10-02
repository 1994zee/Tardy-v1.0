using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tardy.model
{
    class Record_2
    {
        public string client_id { get; set; }
        public string status { get; set; }
        public string company_name { get; set; }
        public string dba { get; set; }
        public string product { get; set; }
        public string pricing_batch_name { get; set; }
        public string surcharge_type { get; set; }
        public string surcharge_description { get; set; }
        public string bill_rates { get; set; }

        public Record_2(string clientID,string Status, string companyName,string DBA,string Product,string pricingBatchName,string surchargeType,
            string surchargeDescription,string billRates)
        {
            this.client_id = clientID;
            this.status = Status;
            this.company_name = companyName;
            this.dba = DBA;
            this.product = Product;
            this.pricing_batch_name = pricingBatchName;
            this.surcharge_type = surchargeType;
            this.surcharge_description = surchargeDescription;
            this.bill_rates = billRates;
        }
    }
}
