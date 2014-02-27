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
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Assigns GUIDs to Excel worksheets.
    /// </summary>
    internal sealed class WorksheetIdentifier
    {
        /// <summary>
        /// Initialises a new WorksheetIdentifier instance.
        /// </summary>
        public WorksheetIdentifier()
        { }

        /// <summary>
        /// Identifies a worksheet whose custom identifier matches the supplied key.
        /// </summary>
        /// <param name="provider">The IWorksheetProvider that provides the eligible worksheets.</param>
        /// <param name="key">The unique GUID string to be matched.</param>
        /// <returns>The matching worksheet, or null if no match could be found.</returns>
        public Excel.Worksheet GetWorksheet(IWorksheetProvider provider, string key)
        {
            Excel.Worksheet[] sheets = provider.GetWorksheets();
            return sheets.FirstOrDefault(s => MatchesKey(s, key));
        }

        /// <summary>
        /// The key used to create and retrieve a worksheet custom property representing a unique value 
        /// assigned to a worksheet by the add-in.
        /// </summary>
        private static readonly string s_cpKey = "06A122BFF0164C5D97B5E0DC51067F0F";

        /// <summary>
        /// Determines whether a worksheet's unique key (if there is one) matches the supplied value.
        /// </summary>
        private static bool MatchesKey(Excel.Worksheet worksheet, string key)
        {
            foreach (Excel.CustomProperty cp in worksheet.CustomProperties) {
                if (cp.Name == s_cpKey && ((string)cp.Value) == key) return true;
            }
            return false;
        }

        /// <summary>
        /// Returns a worksheet's unique key.  If it does not have one then one is created and assigned.
        /// </summary>
        /// <param name="worksheet">The worksheet for which a unique key is required.</param>
        /// <returns>The worksheet's unique string key.</returns>
        public string GetKey(Excel.Worksheet worksheet)
        {
            string key;
            if (!TryGetKey(worksheet, out key)) {
                key = CreateKey(worksheet);
            }
            return key;
        }

        /// <summary>
        /// Returns a value indicating whether a worksheet has been assigned a unique string key.  Outputs
        /// said key if it exists.
        /// </summary>
        private static bool TryGetKey(Excel.Worksheet worksheet, out string key)
        {
            foreach (Excel.CustomProperty cp in worksheet.CustomProperties) {
                if (cp.Name == s_cpKey) {
                    key = (string)cp.Value;
                    return true;
                }
            }
            key = null;
            return false;
        }

        /// <summary>
        /// Creates and assigns a unique string key for a worksheet.
        /// </summary>
        private static string CreateKey(Excel.Worksheet worksheet)
        {
            string key = Guid.NewGuid().ToString("N");
            worksheet.CustomProperties.Add(s_cpKey, key);
            return key;
        }
    }
}