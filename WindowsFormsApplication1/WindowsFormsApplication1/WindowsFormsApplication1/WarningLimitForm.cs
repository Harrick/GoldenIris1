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
    public partial class WarningLimitModify : Form
    {
        public int warningLimit = 0;

        public WarningLimitModify(int limit)
        {
            InitializeComponent();
            this.textBox1.Text = limit.ToString();
        }

        private void clickOK(object sender, EventArgs e)
        {
            try
            {
                if (this.maskedTextBox1.Text == null || this.maskedTextBox1.Text == "")
                {
                    warningLimit = 0;
                }
                else
                {
                    warningLimit = System.Int32.Parse(this.maskedTextBox1.Text);
                }
                this.DialogResult = DialogResult.OK;
            }
            catch(Exception err)
            {
                throw err;
            }
        }

        private void clickCancel(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
