using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System;
using System.Drawing;

namespace SubscribeAndHandleQBEvent
{
    public partial class mainDashboardUI : Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );
        public mainDashboardUI()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void mainDashboardUI_Load(object sender, System.EventArgs e)
        {

        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnDashboard.Height;
            pnlNav.Top = btnDashboard.Top;
            pnlNav.Left = btnDashboard.Left;
            btnDashboard.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnCustomer_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnCustomer.Height;
            pnlNav.Top = btnCustomer.Top;
            btnCustomer.BackColor = Color.FromArgb(46, 51, 73);

            //CustomerForm customerForm = new CustomerForm();
            //customerForm.ShowDialog();
        }

        private void btnEmployee_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnEmployee.Height;
            pnlNav.Top = btnEmployee.Top;
            btnEmployee.BackColor = Color.FromArgb(46, 51, 73);

            //EmployeeForm employeeForm = new EmployeeForm();
            //employeeForm.ShowDialog();
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnReport.Height;
            pnlNav.Top = btnReport.Top;
            btnReport.BackColor = Color.FromArgb(46, 51, 73);

            //GeneralReportForm generalReportForm = new GeneralReportForm();
            //generalReportForm.ShowDialog(); 
        }

        private void btnBill_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnBill.Height;
            pnlNav.Top = btnBill.Top;
            btnBill.BackColor = Color.FromArgb(46, 51, 73);

            //BillForm billForm = new BillForm();
            //billForm.ShowDialog();
        }

        private void btnDashboard_Leave(object sender, EventArgs e)
        {
            btnDashboard.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnCustomer_Leave(object sender, EventArgs e)
        {
            btnCustomer.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnEmployee_Leave(object sender, EventArgs e)
        {
            btnEmployee.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnReport_Leave(object sender, EventArgs e)
        {
            btnReport.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnBill_Leave(object sender, EventArgs e)
        {
            btnBill.BackColor = Color.FromArgb(24, 30, 54);
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
