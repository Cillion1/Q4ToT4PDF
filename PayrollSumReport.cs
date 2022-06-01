using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Q4ToT4PDF
{
    public class PayrollSumReport
    {
        public string employmentIncome { get; set; }
        public string employeeCPPContribution { get; set; }
        public string employerCPPContribution { get; set; }
        public string employeeEIPremium { get; set; }
        public string employerEIPremium { get; set; }
        public string incomeTaxDeducted { get; set; }
        public string totalDeductionsReported { get; set; }
    }
}
