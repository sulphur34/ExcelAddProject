using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddProject
{
    public partial class COKReport : Form
    {
        public COKReport()
        {
            InitializeComponent();
            rbMonth.Checked = true;
            dtpStart.Enabled = false;
            dtpEnd.Enabled = false;
            dtpStart.Value = new DateTime(2018, 1, 1);
            dtpEnd.Value = DateTime.Today;
            dtpMonth.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        }

        private void rbRange_CheckedChanged(object sender, EventArgs e)
        {
            if (rbRange.Checked)
            {
                dtpStart.Enabled = true;
                dtpEnd.Enabled = true;
                dtpMonth.Enabled = false;
            }
            else
            {
                dtpStart.Enabled = false;
                dtpEnd.Enabled = false;
                dtpMonth.Enabled = true;
            }
        }

        private void rbMonth_CheckedChanged(object sender, EventArgs e)
        {
            if (rbMonth.Checked)
            {
                dtpStart.Enabled = false;
                dtpEnd.Enabled = false;
                dtpMonth.Enabled = true;
            }
            else
            {
                dtpStart.Enabled = true;
                dtpEnd.Enabled = true;
                dtpMonth.Enabled = false;
            }
        }

        private void bCountReject_Click(object sender, EventArgs e)
        {
            List<Weld> WBbase = RepairRates.WeldData(cbWBCount.Checked, false);
            if (rbMonth.Checked) RepairRates.PrintratesCOK(RepairRates.CountRatesCOK(dtpStart.Value, dtpMonth.Value.AddMonths(1).AddDays(-1), WBbase, cbDiameter.Checked), RepairRates.CountRatesCOK(dtpMonth.Value, dtpMonth.Value.AddMonths(1).AddDays(-1),
                                 WBbase, cbDiameter.Checked));
            else RepairRates.PrintratesCOK(RepairRates.CountRatesCOK(dtpStart.Value, dtpEnd.Value, WBbase, cbDiameter.Checked), RepairRates.CountRatesCOK(dtpMonth.Value, dtpEnd.Value,
                                 WBbase, cbDiameter.Checked));
        }
    }


}
