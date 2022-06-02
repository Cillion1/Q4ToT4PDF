using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToT4PDF
{
    public class PayrollSumReport
    {
        public string year { get; set; }
        public string numEmployee { get; set; }

        // Box 14
        public string employmentIncome { get; set; }
        // Box 16
        public string employeeCPPContribution { get; set; }
        // Box 27
        public string employerCPPContribution { get; set; }
        // Box 18
        public string employeeEIPremium { get; set; }
        // Box 19
        public string employerEIPremium { get; set; }
        // Box 22
        public string incomeTaxDeducted { get; set; }
        // Box 80 
        // add all the above values up to get the total deduction
        public string totalDeductionsReported { get; set; }

        public PayrollSumReport()
        {
            employmentIncome = "0";
            employeeCPPContribution = "0";
            employerCPPContribution = "0";
            employeeEIPremium = "0";
            employerEIPremium = "0";
            incomeTaxDeducted = "0";
            totalDeductionsReported = "0";
            numEmployee = "0";
        }

        public void calTotalDeduction()
        {
           double temp = Convert.ToDouble(this.incomeTaxDeducted) + Convert.ToDouble(this.employerCPPContribution) + Convert.ToDouble(this.employeeCPPContribution) + Convert.ToDouble(this.employeeEIPremium) + Convert.ToDouble(this.employerEIPremium);
           this.totalDeductionsReported = Convert.ToString(temp);
        }
    }


}
