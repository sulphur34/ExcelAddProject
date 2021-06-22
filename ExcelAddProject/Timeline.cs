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
    public partial class Timeline : Form
    {
        public Timeline()
        {
            InitializeComponent();
            dtpStart.Value = new DateTime(2020, 10, 26);
            dtpEnd.Value = DateTime.Today;
            rbLTCS.Checked = true;
        }
        string Material;
        private void bCountReject_Click(object sender, EventArgs e)
        {
            RepairRates.PrintTimeline(RepairRates.TimelineCount(dtpStart.Value, dtpEnd.Value, Material, cbWBCount.Checked, cbOfficial.Checked));
        }
        private void rbALLOY_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                Material = radioButton.Text;
            }
        }
        private void rbLTCS_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                Material = radioButton.Text;
            }
        }

        private void rbSS_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                Material = radioButton.Text;
            }
        }

        private void rbF22_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                Material = radioButton.Text;
            }
        }
    }
}
