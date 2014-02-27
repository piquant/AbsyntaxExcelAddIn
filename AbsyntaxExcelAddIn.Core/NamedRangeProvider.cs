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
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A provider of details pertaining to range names.
    /// </summary>
    internal sealed class NamedRangeProvider : INamedRangeProvider
    {
        /// <summary>
        /// Initialises a new NamedRangeProvider instance.
        /// </summary>
        /// <param name="provider">An IWorksheetProvider implementation.</param>
        public NamedRangeProvider(IWorksheetProvider provider)
        {
            if (provider == null) {
                throw new ArgumentNullException("provider");
            }
            m_provider = provider;
        }

        /// <summary>
        /// An IWorksheetProvider implementation responsible for supplying the available worksheets on
        /// demand.
        /// </summary>
        private IWorksheetProvider m_provider;

        private Dictionary<string, string> m_worksheetKeys;

        /// <summary>
        /// Returns a populated dictionary of workbook range names and their associated worksheet keys.
        /// </summary>
        private Dictionary<string, string> GetWorksheetKeysDictionary()
        {
            if (m_worksheetKeys == null) {
                m_worksheetKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                Excel.Workbook wb = m_provider.GetWorksheets().First().Application.ActiveWorkbook;
                ObtainWorksheetKeys(m_worksheetKeys, wb);
            }
            return m_worksheetKeys;
        }

        /// <summary>
        /// Populates a dictionary with workbook range names and their associated worksheet keys.
        /// </summary>
        private static void ObtainWorksheetKeys(Dictionary<string, string> dic, Excel.Workbook wb)
        {
            var wi = new WorksheetIdentifier();
            foreach (Excel.Name name in wb.Names) {
                Excel.Worksheet ws = GetWorksheetForRangeName(name);
                if (ws != null) {
                    string rName = name.Name;
                    string wsKey = wi.GetKey(ws);
                    dic[rName] = wsKey;
                }
            }
        }

        /// <summary>
        /// Attempts to find the worksheet associated with a range name.
        /// </summary>
        private static Excel.Worksheet GetWorksheetForRangeName(Excel.Name name)
        {
            Excel.Worksheet ws;
            try {
                Excel.Range r = name.RefersToRange;
                ws = r.Worksheet;
            }
            catch {
                /* This happens when a range name refers to multiple worksheets or contains 
                 * an invalid reference
                 */
                ws = null;
            }
            return ws;
        }

        /// <summary>
        /// Returns a list of known range names.
        /// </summary>
        /// <returns>An array of zero or more range names, ordered alphabetically within range name scope
        /// (worksheet-scoped names are positioned above workbook-scoped names).</returns>
        public string[] GetRangeNames()
        {
            var c = new RangeNameComparer();
            Dictionary<string, string> dic = GetWorksheetKeysDictionary();
            return dic.Select(kvp => kvp.Key).OrderBy(s => s, c).ToArray();
        }

        /// <summary>
        /// Identifies the worksheet associated with a range name.
        /// </summary>
        /// <param name="rangeName">The range name for which the associated worksheet is required.</param>
        /// <returns>The associated worksheet, or null if no association can be found.</returns>
        public Excel.Worksheet IdentifyWorksheet(string rangeName)
        {
            string wsKey;
            Dictionary<string, string> dic = GetWorksheetKeysDictionary();
            if (dic.TryGetValue(rangeName, out wsKey)) {
                var wi = new WorksheetIdentifier();
                return wi.GetWorksheet(m_provider, wsKey);
            }
            return null;
        }

        /// <summary>
        /// Clears this NamedRangeProvider of any cached range name details.
        /// </summary>
        public void Clear()
        {
            m_worksheetKeys = null;
        }

        /// <summary>
        /// A string IComparer implementation for comparing two range names.
        /// </summary>
        /// <remarks>
        /// This comparer ensures that worksheet-scoped range names (i.e. range names containing an
        /// exclamation mark) are considered to be "less than" workbook-scoped range names.
        /// </remarks>
        private sealed class RangeNameComparer : IComparer<string>
        {
            public RangeNameComparer()
            { }

            public int Compare(string x, string y)
            {
                bool bx = x.IndexOf('!') >= 0;
                bool by = y.IndexOf('!') >= 0;
                if ((bx && by) || (!bx && !by)) {
                    return String.Compare(x, y);
                }
                return bx ? -1 : 1;
            }
        }
    }
}