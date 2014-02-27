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
using System.Windows.Forms;
using AbsyntaxExcelAddIn.Resources;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A Form that allows users to select the type of Absyntax licence that the add-in should use.
    /// </summary>
    /// <remarks>
    /// The options are a full licence or a third-party execution licence.  The latter requires a unique 
    /// client identifier to be specified and this identifier must match the identifier assigned during 
    /// creation of the associated licence key.
    /// </remarks>
    public partial class LicenceDialogue : Form
    {
        public LicenceDialogue()
        {
            InitializeComponent();
            Text = TextResources.Title_LicenceDialogue;
            DialogResult = DialogResult.Cancel;
            LicenceDialogueContent content = Content;
            content.Accepted += Content_Accepted;
            content.Cancelled += Content_Cancelled;
        }

        private LicenceDialogueContent Content
        {
            get { return (LicenceDialogueContent)m_elementHost.Child; }
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
        /// Gets or sets a value indicating whether a full Absyntax licence is to be used.
        /// </summary>
        /// <remarks>
        /// If this property is false, the implication is that a third-party execution licence is to be used.
        /// </remarks>
        public bool UsesFullLicence
        {
            get { return Content.UsesFullLicence; }
            set { Content.UsesFullLicence = value; }
        }

        /// <summary>
        /// Gets or sets a nullable Guid representing the client identifier to be used in conjunction with a 
        /// third-party licence.
        /// </summary>
        public Guid? ClientId
        {
            get {
                Guid g;
                return Guid.TryParse(Content.ClientId, out g) ? g : (Guid?)null;
            }
            set {
                string cid;
                if (value.HasValue) {
                    cid = value.Value.ToString("N");
                }
                else {
                    cid = null;
                }
                Content.ClientId = cid;
            }
        }
    }
}