/* Copyright © 2013 Managing Infrastructure Information Ltd
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without modification, are permitted provided 
 * that the following conditions are met:
 * 
 * 1. Redistributions of source code must retain the above copyright notice, this list of conditions and the 
 * following disclaimer.
 * 
 * 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and 
 * the following disclaimer in the documentation and/or other materials provided with the distribution.
 * 
 * 3. Neither the name Managing Infrastructure Information Ltd (MIIL) nor the names of its contributors may 
 * be used to endorse or promote products derived from this software without specific prior written 
 * permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED 
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A 
 * PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR 
 * ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT 
 * LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR 
 * TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * */

using AbsyntaxExcelAddIn.Properties;
using System.Configuration;

namespace AbsyntaxExcelAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.m_absyntaxGroup = this.Factory.CreateRibbonGroup();
            this.m_configure = this.Factory.CreateRibbonButton();
            this.m_run = this.Factory.CreateRibbonButton();
            this.m_licence = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.m_absyntaxGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabData";
            this.tab1.Groups.Add(this.m_absyntaxGroup);
            this.tab1.Label = "TabData";
            this.tab1.Name = "tab1";
            // 
            // m_absyntaxGroup
            // 
            this.m_absyntaxGroup.Items.Add(this.m_configure);
            this.m_absyntaxGroup.Items.Add(this.m_run);
            this.m_absyntaxGroup.Items.Add(this.m_licence);
            this.m_absyntaxGroup.Label = "Absyntax";
            this.m_absyntaxGroup.Name = "m_absyntaxGroup";
            // 
            // m_configure
            // 
            this.m_configure.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.m_configure.Image = ((System.Drawing.Image)(resources.GetObject("m_configure.Image")));
            this.m_configure.Label = "Configure";
            this.m_configure.Name = "m_configure";
            this.m_configure.ShowImage = true;
            this.m_configure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Configure_Click);
            // 
            // m_run
            // 
            this.m_run.Enabled = false;
            this.m_run.Image = ((System.Drawing.Image)(resources.GetObject("m_run.Image")));
            this.m_run.Label = "Run";
            this.m_run.Name = "m_run";
            this.m_run.ShowImage = true;
            this.m_run.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Run_Click);
            // 
            // m_licence
            // 
            this.m_licence.Image = ((System.Drawing.Image)(resources.GetObject("m_licence.Image")));
            this.m_licence.Label = "Set Licence";
            this.m_licence.Name = "m_licence";
            this.m_licence.ShowImage = true;
            this.m_licence.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetLicence_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.m_absyntaxGroup.ResumeLayout(false);
            this.m_absyntaxGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton m_configure;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup m_absyntaxGroup;
        private Microsoft.Office.Tools.Ribbon.RibbonButton m_run;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton m_licence;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
