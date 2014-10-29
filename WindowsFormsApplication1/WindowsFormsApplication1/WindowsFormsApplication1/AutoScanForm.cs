using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplicationAutoScan
{
    public partial class AutoScanForm : Form
    {
        int warningLimit = 0;
        int scanHour = 0;

        public AutoScanForm()
        {
            InitializeComponent();
            this.Log.Text = "Welcome to use ScanStock!" + "\n";
        }

        private void browseLoadButtonClick(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            if (DialogResult.OK == openDialog.ShowDialog())
            {
                this.ReportPath.Text = openDialog.FileName;
            }
        }

        private void browseScanButtonClick(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            if (DialogResult.OK == openDialog.ShowDialog())
            {
                this.ScanPath.Text = openDialog.FileName;
            }
        }

        private void setWarningLimit(object sender, EventArgs e)
        {
            WarningLimitModify limitForm = new WarningLimitModify(warningLimit);
            if (DialogResult.OK == limitForm.ShowDialog())
            {
                warningLimit = limitForm.warningLimit;
            }
        }

        private void hideMethod(object sender, EventArgs e)
        {
            this.Hide();
            this.notifyIcon1.Visible = true;
        }

        private void notifyIconClick(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.notifyIcon1.Visible = false;
        }

        /*
        private void loadMethod(object sender, EventArgs e)
        {
            Stock stock = new Stock(null,ReportPath.Text,warningLimit,this.Log);
            stock.loadScan();
        }*/

        private void scanMethod(object sender, EventArgs e)
        {
            Stock stock = new Stock(ScanPath.Text, ReportPath.Text, warningLimit, this.Log);
            stock.scanStock();
        }

        private void clearLogClick(object sender, EventArgs e)
        {
            this.Log.Text = "";
        }

        private void saveLogClick(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            if (DialogResult.OK == openDialog.ShowDialog())
            {
                StreamWriter writer = new StreamWriter(openDialog.FileName);
                writer.Write(this.Log.Text);
                writer.Close();
            }
        }

        private void versionClick(object sender, EventArgs e)
        {
            MessageBox.Show("V1.0", "Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void timerTick(object sender, EventArgs e)
        {
            this.Log.AppendText("Timer out! " + DateTime.Now.Hour + " \n");
            if (DateTime.Now.Hour == scanHour)
            {
                // this.loadMethod(sender, e);
                this.scanMethod(sender, e);
            }
        }

        private void setAutoScan(object sender, EventArgs e)
        {
            setAutoScanForm timeForm = new setAutoScanForm(scanHour, this.timer1.Enabled);
            if (DialogResult.OK == timeForm.ShowDialog())
            {
                scanHour = timeForm.hour;
                this.timer1.Enabled = timeForm.enableAutoScan.Checked;
            }
        }
    }
}
