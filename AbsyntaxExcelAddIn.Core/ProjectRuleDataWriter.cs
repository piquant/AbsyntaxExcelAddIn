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
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// An IDataWriter implementation that write project invocation rules to a worksheet.
    /// </summary>
    internal sealed class ProjectRuleDataWriter : IDataWriter
    {
        /// <summary>
        /// Initialises a new ProjectRuleDataWriter instance.
        /// </summary>
        public ProjectRuleDataWriter()
        { }

        /// <summary>
        /// Writes the various field values of a set of ProjectInvocationRules to a worksheet.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet to which project invocation rule data is to be 
        /// written.</param>
        /// <param name="rules">The set of ProjectInvocationRules to be written.</param>
        /// <param name="firstRow">The 1-based worksheet row index at which to start writing project 
        /// invocation rule data.</param>
        public void Write(Excel.Worksheet worksheet, ProjectInvocationRule[] rules, int firstRow)
        {
            m_worksheet = worksheet;
            m_row = firstRow;
            foreach (ProjectInvocationRule rule in rules) {
                m_colIndex = 1;
                rule.Write(this);
                m_row++;
            }
        }

        private Excel.Worksheet m_worksheet;

        /// <summary>
        /// Tracks the current 1-based column index at which the next data value will be written.
        /// </summary>
        private int m_colIndex;

        /// <summary>
        /// Tracks the current 1-based row index at which the next data value will be written.
        /// </summary>
        private int m_row;

        /// <summary>
        /// Writes an object of type T to the current row and column.
        /// </summary>
        /// <param name="value">The T to be written.</param>
        public void Write<T>(T value)
        {
            Excel.Worksheet ws = m_worksheet;
            if (ws == null) {
                throw new InvalidOperationException();
            }
            Helper.WriteCell(ws, m_colIndex++, m_row, value);
        }
    }
}