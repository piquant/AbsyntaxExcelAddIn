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

using System;
using System.Linq;
using System.Windows.Forms;
using AbsyntaxExcelAddIn.Resources;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A Form supporting a list of Absyntax project invocation rules associated with the active workbook.
    /// </summary>
    public partial class ProjectConfigurationDialogue : Form
    {
        public ProjectConfigurationDialogue(IWorksheetProvider provider)
        {
            InitializeComponent();
            Text = TextResources.Title_ProjectConfigurationDialogue;
            m_provider = provider;
            DialogResult = DialogResult.Cancel;
            ConfigurationDialogueContent content = Content;
            content.WorksheetProvider = provider;
            content.Accepted += Content_Accepted;
            content.Cancelled += Content_Cancelled;
        }

        private IWorksheetProvider m_provider;

        private ConfigurationDialogueContent Content
        {
            get { return (ConfigurationDialogueContent)m_elementHost.Child; }
        }

        private void Content_Accepted(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Content_Cancelled(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Gets or sets the execution mode to be adopted when invoking the set of rules.
        /// </summary>
        public ExecutionMode Mode
        {
            get { return Content.Mode; }
            set { Content.Mode = value; }
        }

        /// <summary>
        /// Sets this ProjectConfigurationDialogue's collection of project invocation rules.
        /// </summary>
        /// <remarks>
        /// The supplied rules are cloned.  This is because they are used as view-models in this dialogue
        /// and thus the user can change their states.  If, subsequently, the user cancels this dialogue 
        /// then any changes to individual ProjectInvocationRule instances can be discarded.
        /// </remarks>
        public void SetRules(ProjectInvocationRule[] rules)
        {
            rules = GetClones(rules);
            ConfigurationDialogueContent content = Content;
            content.SetRules(rules);
        }

        /// <summary>
        /// Returns clones of this ProjectConfigurationDialogue's collection of project invocation rules.
        /// </summary>
        public ProjectInvocationRule[] GetRules()
        {
            return Content.GetRules();
        }

        /// <summary>
        /// Clones the supplied ProjectInvocationRule instances.
        /// </summary>
        private static ProjectInvocationRule[] GetClones(ProjectInvocationRule[] rules)
        {
            return rules == null ? null : rules.Select(r => r.Clone()).ToArray();
        }
    }
}