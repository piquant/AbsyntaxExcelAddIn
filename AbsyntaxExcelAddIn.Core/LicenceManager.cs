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
using AbsyntaxExcelAddIn.Core.Properties;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Interacts with roaming-user isolated storage settings in order to read and write Absyntax licence-
    /// related details.
    /// </summary>
    public sealed class LicenceManager
    {
        public LicenceManager()
        { }

        /// <summary>
        /// Acquires the current Absyntax client id to be used in conjunction with an associated third-party 
        /// execution licence.
        /// </summary>
        /// <remarks>
        /// If the current seat (i.e. user/machine combination) has a full, activated and unexpired licence 
        /// then this is enough to allow this Excel add-in to work in conjunction with the Absyntax runtime.  
        /// Equally an activated, unexpired third-party execution licence may be used as long as the Absyntax 
        /// API is presented with a compatible client identifier (i.e. a GUID created in conjunction with the 
        /// licence key).
        /// </remarks>
        public Guid? GetClientId()
        {
            string value = Settings.Default.ClientId;
            Guid? guid;
            if (String.IsNullOrEmpty(value)) {
                guid = null;
            }
            else {
                Guid g;
                if (Guid.TryParse(value, out g)) {
                    guid = g;
                }
                else {
                    guid = null;
                }
            }
            return guid;
        }

        /// <summary>
        /// Stores the supplied nullable Guid client id.
        /// </summary>
        public void SetClientId(Guid? guid)
        {
            Settings s = Settings.Default;
            s.ClientId = guid == null ? null : guid.Value.ToString();
            s.Save();
        }

        /// <summary>
        /// Gets or sets a value indicating whether a full Absyntax licence is to be targeted.
        /// </summary>
        public bool UsesFullLicence
        {
            get { return Settings.Default.UsesFullLicence; }
            set {
                Settings s = Settings.Default;
                s.UsesFullLicence = value;
                s.Save();
            }
        }
    }
}