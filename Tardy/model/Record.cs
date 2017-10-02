using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tardy.model
{
    class Record
    {
        public int days_since_reported { get; set; }
        public string carrier_claim_number { get; set; }
        public string cs_claim_number { get; set; }
        public string bill_date { get; set; }
        public string bill_event_code { get; set; }
        public string bill_unit { get; set; }
        public string client_id { get; set; }
        public string employee_id { get; set; }
        public string injured_employee { get; set; }
        public string location { get; set; }
        public string comment { get; set; }
        public string doi { get; set; }
        public string bill_rate { get; set; }
        public Record(string daysSinceReported,string carrierClaimNumber,string csClaimNumber,string billDate,string billEventCode,string billUnit
            ,string clientId,string employeeId,string injuredEmployee,string Location,string Comment,string DOI)
        {
            if ( daysSinceReported != "")
            {
                this.days_since_reported = Convert.ToInt32(daysSinceReported);
            }
            else
            {
                this.days_since_reported = 0;
            }
            this.carrier_claim_number = carrierClaimNumber;
            this.cs_claim_number = csClaimNumber;
            this.bill_date = billDate;
            this.bill_event_code = billEventCode;
            this.bill_unit = billUnit;
            this.client_id = clientId;
            this.employee_id = employeeId;
            this.injured_employee = injuredEmployee;
            this.location = Location;
            this.comment = Comment;
            this.doi = DOI;
        }
    }
}
