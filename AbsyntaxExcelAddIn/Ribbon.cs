/* Copyright © 2013-2014 Managing Infrastructure Information Ltd
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

using System;
using System.Configuration;
using System.Drawing;
using System.Windows.Forms;
using AbsyntaxExcelAddIn.Core;
using AbsyntaxExcelAddIn.Resources;
using Microsoft.Office.Tools.Ribbon;

namespace AbsyntaxExcelAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            m_configure.Label = TextResources.Label_Configure;
            m_configure.ScreenTip = TextResources.ScreenTip_Configure;
            m_configure.SuperTip = TextResources.SuperTip_Configure;
            m_run.Label = TextResources.Label_Run;
            m_run.ScreenTip = TextResources.ScreenTip_Run;
            m_licence.Label = TextResources.Label_SetLicence;
            m_licence.ScreenTip = TextResources.ScreenTip_SetLicence;
            m_licence.SuperTip = TextResources.SuperTip_SetLicence;
            m_licence.Enabled = ConfigAllowsLicenceAccess();
            
            ThisAddIn addIn = Globals.ThisAddIn;
            addIn.HostStatusChanged += AddIn_HostStatusChanged;
            addIn.StateChanged += AddIn_StateChanged;
            SetButtonStates();
            UpdateRunButtonSuperTip();
        }

        private static bool ConfigAllowsLicenceAccess()
        {
            string s = ConfigurationManager.AppSettings["AllowLicenceAccess"];
            bool value;
            return bool.TryParse(s, out value) ? value : false;
        }

        private void AddIn_HostStatusChanged(object sender, EventArgs e)
        {
            UpdateRunButtonSuperTip();
        }

        private void UpdateRunButtonSuperTip()
        {
            ThisAddIn addIn = Globals.ThisAddIn;
            string tip = TextResources.SuperTip_Run;
            if (!addIn.ProcessAvailable) {
                Helper.AddParagraph(ref tip, TextResources.Msg_HostProcessNotOpen);
            }
            else if (!addIn.ServiceAvailable) {
                Helper.AddParagraph(ref tip, TextResources.Msg_HostServiceNotAvailable);
            }
            else if (addIn.ExecutionState != AddInExecutionState.CanExecute) {
                Helper.AddParagraph(ref tip, TextResources.Msg_NoValidRules);
            }
            m_run.SuperTip = tip;
        }

        private void AddIn_StateChanged(object sender, System.EventArgs e)
        {
            SetButtonStates();
            UpdateRunButtonSuperTip();
        }

        private void SetButtonStates()
        {
            ThisAddIn addIn = Globals.ThisAddIn;
            m_configure.Enabled = addIn.HasActiveWorkbook && addIn.ExecutionState != AddInExecutionState.Executing;
            m_run.Enabled = addIn.ExecutionState == AddInExecutionState.CanExecute;
        }

        /// <summary>
        /// The last bounds of the ProjectConfigurationDialogue.
        /// </summary>
        private Rectangle? m_bounds;

        private void Configure_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn addIn = Globals.ThisAddIn;
            var d = new ProjectConfigurationDialogue(addIn);
            if (m_bounds.HasValue) {
                d.Bounds = m_bounds.Value;
                d.StartPosition = FormStartPosition.Manual;
            }
            d.Mode = addIn.Mode;
            ProjectInvocationRule[] rules = addIn.Rules;
            d.SetRules(rules);
            var result = d.ShowDialog();
            if (result == DialogResult.OK) {
                addIn.Mode = d.Mode;
                addIn.Rules = rules = d.GetRules();
                addIn.WriteRules();
            }
            m_bounds = d.Bounds;
        }

        private void Run_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn addIn = Globals.ThisAddIn;
            addIn.Run();
        }

        private void SetLicence_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn addIn = Globals.ThisAddIn;
            bool full = addIn.UsesFullLicence;
            Guid? g = addIn.GetClientId();
            var f = new LicenceDialogue();
            f.UsesFullLicence = full;
            f.ClientId = g;
            var result = f.ShowDialog();
            if (result == DialogResult.OK) {
                addIn.ChangeLicenceDetails(f.UsesFullLicence, f.ClientId);
            }
        }
    }
}