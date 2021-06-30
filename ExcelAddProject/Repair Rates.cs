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
    public partial class RepairForm : Form
    {
        public RepairForm()
        {
            InitializeComponent();
            rbAll.Checked = true;
            dateTimePicker2.Value = new DateTime(2018, 1, 1);
            dateTimePicker1.Value = DateTime.Today;

        }
        string ProdObject = "All";
        string DataSource = "Repair Rates";
        bool WBcalc = false;
        private void button1_Click(object sender, EventArgs e)
        {
            if (cbDiameter.Checked == false & cbNoDivision.Checked == false)
            {
                DataSource = "Repair Rates";
            }
            else if (cbDiameter.Checked == true & cbNoDivision.Checked == false)
            {
                DataSource = "Simple Rates";
            }
            else if (cbDiameter.Checked == false & cbNoDivision.Checked == true) 
            {
                DataSource = "Disqual Rates";
            }
            this.WindowState = FormWindowState.Minimized;
            RepairRates.PrintRates(dateTimePicker2.Value, dateTimePicker1.Value, ProdObject, WBcalc, cbOfficial.Checked, DataSource, cbRepairValid.Checked, cbVolume.Checked);
            this.WindowState = FormWindowState.Normal;
        }

        private void rbAll_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                ProdObject = "All";
            }

        }

        private void rbErection_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                ProdObject = "Erection";
            }
        }

        private void rbWorkshop_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                ProdObject = "Workshop";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                WBcalc = true;
            else
                WBcalc = false;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbRepairValid_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbDiameter_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDiameter.Checked)
            {
                cbNoDivision.Checked = false;
                cbNoDivision.Enabled = false;            
            }
            else
            {
                cbNoDivision.Checked = false;
                cbNoDivision.Enabled = true;
            }
        }

        private void cbNoDivision_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNoDivision.Checked)
            {
                cbDiameter.Checked = false;
                cbDiameter.Enabled = false;
            }
            else
            {
                cbDiameter.Checked = false;
                cbDiameter.Enabled = true;
            }
        }
    }
}
