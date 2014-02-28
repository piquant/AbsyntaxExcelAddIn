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
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// An IDataReader implementation that reads project invocation rules from a worksheet.
    /// </summary>
    internal sealed class ProjectRuleDataReader : IDataReader
    {
        /// <summary>
        /// Initialises a new ProjectRuleDataReader instance.
        /// </summary>
        /// <param name="wsProvider">An IWorksheetProvider implementation.</param>
        /// <param name="nrProvider">An INamedRangeProvider implementation</param>
        public ProjectRuleDataReader(IWorksheetProvider wsProvider, INamedRangeProvider nrProvider)
        {
            m_wsProvider = wsProvider;
            m_nrProvider = nrProvider;
        }

        private IWorksheetProvider m_wsProvider;

        private INamedRangeProvider m_nrProvider;

        /// <summary>
        /// Converts worksheet data into a collection of zero or more ProjectInvocationRules.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet in which project invocation rule data is to be 
        /// found.</param>
        /// <param name="firstRow">The 1-based worksheet row index from which to start searching for
        /// project invocation rule data.</param>
        /// <returns>An array of zero or more ProjectInvocationRule instances.</returns>
        public ProjectInvocationRule[] Read(Excel.Worksheet worksheet, int firstRow)
        {
            m_worksheet = worksheet;
            m_row = firstRow;
            var list = new List<ProjectInvocationRule>();
            while (RowHasData(worksheet, m_row)) {
                m_colIndex = 1;
                try {
                    var rule = new ProjectInvocationRule(m_wsProvider, m_nrProvider, this);
                    list.Add(rule);
                }
                catch { }
                m_row++;
            }
            return list.ToArray();
        }

        /// <summary>
        /// Returns a value indicating whether there is data in a worksheet row.
        /// </summary>
        private static bool RowHasData(Excel.Worksheet ws, int row)
        {
            string cell = Helper.GetCell("A", row);
            return ws.Range[cell].Value2 != null;
        }

        private Excel.Worksheet m_worksheet;

        /// <summary>
        /// Tracks the current 1-based column index from which the next data value will be read.
        /// </summary>
        private int m_colIndex;

        /// <summary>
        /// Tracks the current 1-based row index from which the next data value will be read.
        /// </summary>
        private int m_row;

        /// <summary>
        /// Returns an object of type T from the current row and column.
        /// </summary>
        /// <typeparam name="T">The type of object to be returned.</typeparam>
        /// <returns>An object of type T.</returns>
        public T Read<T>()
        {
            Excel.Worksheet ws = m_worksheet;
            if (ws == null) {
                throw new InvalidOperationException();
            }
            return Helper.ReadCell<T>(ws, m_colIndex++, m_row);
        }
    }
}