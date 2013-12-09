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

using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Coordinates the writing and reading of project invocation rules to and from an Excel worksheet.
    /// </summary>
    public sealed class PersistedRuleManager
    {
        /// <summary>
        /// Initialises a new PersistedRuleManager instance.
        /// </summary>
        public PersistedRuleManager()
        { }

        /// <summary>
        /// Writes a set of project invocation rules to a worksheet.
        /// </summary>
        /// <param name="ws">The worksheet to which the rules are to be written.</param>
        /// <param name="mode">The ExecutionMode to be written.</param>
        /// <param name="rules">The project invocation rules to be written.</param>
        public void Save(Excel.Worksheet ws, ExecutionMode mode, ProjectInvocationRule[] rules)
        {
            ws.Cells.ClearContents();
            Helper.WriteCell(ws, 1, 1, mode.ToString());
            var writer = new ProjectRuleDataWriter();
            writer.Write(ws, rules, 2);
        }

        /// <summary>
        /// Reads a set of project invocation rules from a worksheet.
        /// </summary>
        /// <remarks>
        /// The worksheet is expected to store the execution mode in cell A1.  Each invocation rule occupies 
        /// its own row, starting at row 2: each column contains a field value and all such values are used 
        /// collectively to rehydrate a rule.
        /// </remarks>
        /// <param name="provider">An IWorksheetProvider implementation.</param>
        /// <param name="ws">The worksheet containing the rules to be read.</param>
        /// <param name="mode">The ExecutionMode read from the worksheet.</param>
        /// <param name="rules">The project invocation rules read from the worksheet.</param>
        public void Load(IWorksheetProvider provider, Excel.Worksheet ws, out ExecutionMode mode, out ProjectInvocationRule[] rules)
        {
            mode = Helper.ReadCell<ExecutionMode>(ws, 1, 1);
            var reader = new ProjectRuleDataReader(provider);
            rules = reader.Read(ws, 2);
        }
    }
}