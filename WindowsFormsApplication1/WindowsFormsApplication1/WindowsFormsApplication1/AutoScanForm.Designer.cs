namespace WindowsFormsApplicationAutoScan
{
    partial class AutoScanForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoScanForm));
            this.ReportPath = new System.Windows.Forms.TextBox();
            this.BrowseLoad = new System.Windows.Forms.Button();
            this.Load = new System.Windows.Forms.Button();
            this.ScanPath = new System.Windows.Forms.TextBox();
            this.BrowseScan = new System.Windows.Forms.Button();
            this.scan = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Log = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.warningLimitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoScanToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.versionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Save = new System.Windows.Forms.Button();
            this.Clear = new System.Windows.Forms.Button();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ReportPath
            // 
            this.ReportPath.AccessibleName = "";
            this.ReportPath.Location = new System.Drawing.Point(126, 92);
            this.ReportPath.Name = "ReportPath";
            this.ReportPath.Size = new System.Drawing.Size(227, 20);
            this.ReportPath.TabIndex = 0;
            // 
            // BrowseLoad
            // 
            this.BrowseLoad.AccessibleName = "";
            this.BrowseLoad.Location = new System.Drawing.Point(359, 89);
            this.BrowseLoad.Name = "BrowseLoad";
            this.BrowseLoad.Size = new System.Drawing.Size(43, 23);
            this.BrowseLoad.TabIndex = 1;
            this.BrowseLoad.Text = "...";
            this.BrowseLoad.UseVisualStyleBackColor = true;
            this.BrowseLoad.Click += new System.EventHandler(this.browseLoadButtonClick);
            // 
            // Load
            // 
            this.Load.AccessibleName = "";
            this.Load.Location = new System.Drawing.Point(424, 89);
            this.Load.Name = "Load";
            this.Load.Size = new System.Drawing.Size(75, 23);
            this.Load.TabIndex = 2;
            this.Load.Text = "LOAD";
            this.Load.UseVisualStyleBackColor = true;
            // 
            // ScanPath
            // 
            this.ScanPath.AccessibleName = "";
            this.ScanPath.Location = new System.Drawing.Point(126, 45);
            this.ScanPath.Name = "ScanPath";
            this.ScanPath.Size = new System.Drawing.Size(226, 20);
            this.ScanPath.TabIndex = 3;
            // 
            // BrowseScan
            // 
            this.BrowseScan.AccessibleName = "";
            this.BrowseScan.Location = new System.Drawing.Point(359, 42);
            this.BrowseScan.Name = "BrowseScan";
            this.BrowseScan.Size = new System.Drawing.Size(44, 23);
            this.BrowseScan.TabIndex = 4;
            this.BrowseScan.Text = "...";
            this.BrowseScan.UseVisualStyleBackColor = true;
            this.BrowseScan.Click += new System.EventHandler(this.browseScanButtonClick);
            // 
            // scan
            // 
            this.scan.AccessibleName = "";
            this.scan.Location = new System.Drawing.Point(424, 42);
            this.scan.Name = "scan";
            this.scan.Size = new System.Drawing.Size(75, 23);
            this.scan.TabIndex = 5;
            this.scan.Text = "SCAN";
            this.scan.UseVisualStyleBackColor = true;
            this.scan.Click += new System.EventHandler(this.scanMethod);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Report File Path";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Scan File Path";
            // 
            // Log
            // 
            this.Log.AccessibleName = "";
            this.Log.Location = new System.Drawing.Point(15, 157);
            this.Log.Multiline = true;
            this.Log.Name = "Log";
            this.Log.ReadOnly = true;
            this.Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Log.Size = new System.Drawing.Size(469, 289);
            this.Log.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 141);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Log";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.optionsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.menuStrip1.Size = new System.Drawing.Size(511, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.warningLimitToolStripMenuItem,
            this.autoScanToolStripMenuItem,
            this.hideToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.optionsToolStripMenuItem.Text = "Options";
            // 
            // warningLimitToolStripMenuItem
            // 
            this.warningLimitToolStripMenuItem.Name = "warningLimitToolStripMenuItem";
            this.warningLimitToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.warningLimitToolStripMenuItem.Text = "Set warning limit";
            this.warningLimitToolStripMenuItem.Click += new System.EventHandler(this.setWarningLimit);
            // 
            // autoScanToolStripMenuItem
            // 
            this.autoScanToolStripMenuItem.Name = "autoScanToolStripMenuItem";
            this.autoScanToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.autoScanToolStripMenuItem.Text = "Set auto scan";
            this.autoScanToolStripMenuItem.Click += new System.EventHandler(this.setAutoScan);
            // 
            // hideToolStripMenuItem
            // 
            this.hideToolStripMenuItem.Name = "hideToolStripMenuItem";
            this.hideToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.hideToolStripMenuItem.Text = "Hide";
            this.hideToolStripMenuItem.Click += new System.EventHandler(this.hideMethod);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.versionToolStripMenuItem});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.aboutToolStripMenuItem.Text = "About";
            // 
            // versionToolStripMenuItem
            // 
            this.versionToolStripMenuItem.Name = "versionToolStripMenuItem";
            this.versionToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.versionToolStripMenuItem.Text = "Version";
            this.versionToolStripMenuItem.Click += new System.EventHandler(this.versionClick);
            // 
            // Save
            // 
            this.Save.Location = new System.Drawing.Point(255, 467);
            this.Save.Name = "Save";
            this.Save.Size = new System.Drawing.Size(75, 23);
            this.Save.TabIndex = 11;
            this.Save.Text = "Save";
            this.Save.UseVisualStyleBackColor = true;
            this.Save.Click += new System.EventHandler(this.saveLogClick);
            // 
            // Clear
            // 
            this.Clear.Location = new System.Drawing.Point(366, 467);
            this.Clear.Name = "Clear";
            this.Clear.Size = new System.Drawing.Size(75, 23);
            this.Clear.TabIndex = 12;
            this.Clear.Text = "Clear";
            this.Clear.UseVisualStyleBackColor = true;
            this.Clear.Click += new System.EventHandler(this.clearLogClick);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.Click += new System.EventHandler(this.notifyIconClick);
            // 
            // timer1
            // 
            this.timer1.Interval = 3600000;
            this.timer1.Tick += new System.EventHandler(this.timerTick);
            // 
            // AutoScanForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(511, 534);
            this.Controls.Add(this.Clear);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Log);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.scan);
            this.Controls.Add(this.BrowseScan);
            this.Controls.Add(this.ScanPath);
            this.Controls.Add(this.Load);
            this.Controls.Add(this.BrowseLoad);
            this.Controls.Add(this.ReportPath);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "AutoScanForm";
            this.Text = "Golden Iris";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BrowseLoad;
        private System.Windows.Forms.Button Load;
        private System.Windows.Forms.Button BrowseScan;
        private System.Windows.Forms.Button scan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem warningLimitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoScanToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem hideToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem versionToolStripMenuItem;
        private System.Windows.Forms.Button Save;
        private System.Windows.Forms.Button Clear;
        public System.Windows.Forms.TextBox ReportPath;
        public System.Windows.Forms.TextBox ScanPath;
        public System.Windows.Forms.TextBox Log;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.Timer timer1;
    }
}

