using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DavisInvoice
{
    public partial class PostMonth : Form
    {

        public string StrPostMonth { get; set; }

        public PostMonth()
        {
            InitializeComponent();

            postMonthData.Format = DateTimePickerFormat.Custom;
            postMonthData.CustomFormat = "MM/yyyy";
        }

        private void PostMonth_Load(object sender, EventArgs e)
        {
            
        }

        private void Submit_Click(object sender, EventArgs e)
        {
            string strPostMonth = postMonthData.Value.Month.ToString() + "/" + postMonthData.Value.Year.ToString();
            this.StrPostMonth = strPostMonth;
            this.Close();
        }

        private void postMonthData_ValueChanged(object sender, EventArgs e)
        {
            
        }
    }
}
