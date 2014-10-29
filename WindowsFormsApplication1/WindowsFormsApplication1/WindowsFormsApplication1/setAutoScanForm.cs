using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplicationAutoScan
{
    public partial class setAutoScanForm : Form
    {
        public int hour;

        public setAutoScanForm(int scanHour, bool enableAutoScanTimer)
        {
            InitializeComponent();
            this.textBox1.Text = scanHour.ToString();
            hour = scanHour;
            this.enableAutoScan.Checked = enableAutoScanTimer;
        }

        private void buttonnOKClick(object sender, EventArgs e)
        {
            try
            {
                if (this.maskedTextBox1.Text != null && this.maskedTextBox1.Text != "")
                {
                    hour = System.Int32.Parse(this.maskedTextBox1.Text);
                }
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception err)
            {
                throw err;
            }
        }

        private void buttonCancelClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
