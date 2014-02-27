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
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Defines one or more ranges of contiguous row/column cells using a string notation that specifies 
    /// either a named range, a single cell or the cells of any pair of diagonally opposed corners.
    /// </summary>
    /// <remarks>
    /// Excepting range names, the supported notations are "CaRb" and "CaRb:CcRd".  C is a column label, R 
    /// is a row number.  A colon is used to separate the two cell descriptors where a range consisting of 
    /// more than one cell is to be specified.
    /// <para />
    /// Examples include:
    /// <para />
    /// "A1"            A single cell at A1.
    /// "A1:A1"         A single cell at A1.
    /// "C10:E15"       A range of 18 cells, the top-left at C10 and the bottom-right at E15.
    /// "Z20:AA1"       A range of 40 cells, the top-left at Z1 and the bottom-right at AA20.
    /// "BB10:BA10"     A range of two cells, the top-left at BA10 and the bottom-right at BB10.
    /// "MyRange"       A range of cells defined by the specified, workbook-scoped name.  The cells need 
    ///                 not be contiguous (i.e. the range may consist of multiple areas) but may only belong 
    ///                 to one worksheet.
    /// "Sheet1!Local"  A range of cells defined by the specified, worksheet-scoped name.  The cells need 
    ///                 not be contiguous (i.e. the range may consist of multiple areas).  Note that the range 
    ///                 name includes the worksheet by which it is scoped, but the associated areas need not 
    ///                 be on the same worksheet.  They must, however, belong to a single worksheet.
    /// </remarks>
    internal sealed class CellRangeValidator
    {
        /// <summary>
        /// Initialises a new CellRangeValidator instance.
        /// </summary>
        /// <param name="range">A string representation that specifies the cells of any pair of diagonally 
        /// opposed corners of the range being represented.</param>
        /// <param name="provider">An INamedRangeProvider implementation.</param>
        public CellRangeValidator(string range, INamedRangeProvider provider)
        {
            m_range = range;
            SetState(provider);
        }

        /// <summary>
        /// A regular expression used to validate single-cell string notations.
        /// </summary>
        private static readonly Regex s_cellRegex = new Regex("^[A-Z]+[1-9][0-9]*$");
        
        /// <summary>
        /// A regular expression used to validate cell range string notations.
        /// </summary>
        private static readonly Regex s_rangeRegex = new Regex("^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$");

        private string m_range;

        /// <summary>
        /// Gets the cell range notation supplied during instantiation, formatted according to 
        /// its type.
        /// </summary>
        public string Range
        {
            get { return m_range; }
        }

        /// <summary>
        /// Sets the state of this CellRangeValidator based on the range notation supplied during 
        /// instantiation.
        /// </summary>
        private void SetState(INamedRangeProvider provider)
        {
            if (provider.IdentifyWorksheet(m_range) != null) {
                IsValid = true;
            }
            else {
                bool match;
                string r = m_range.ToUpper();
                string firstCell = null, lastCell = null;
                match = s_cellRegex.IsMatch(r);
                if (match) {
                    firstCell = lastCell = r;
                }
                else {
                    match = s_rangeRegex.IsMatch(r);
                    if (match) {
                        int pos = r.IndexOf(":");
                        firstCell = r.Substring(0, pos);
                        lastCell = r.Substring(pos + 1, r.Length - pos - 1);
                    }
                }
                IsValid = match ? CellIsInRange(firstCell) && CellIsInRange(lastCell) : false;
            }
        }

        private static readonly char[] s_digits = "123456789".ToCharArray(); // Don't need zero

        /// <summary>
        /// Determines whether a cell descriptor represents a valid Excel worksheet cell.
        /// </summary>
        private static bool CellIsInRange(string cell)
        {
            int pos = cell.IndexOfAny(s_digits);
            string col = cell.Substring(0, pos);
            int row = Int32.Parse(cell.Substring(pos, cell.Length - pos));
            return CellIsInRange(row, col);
        }

        /// <summary>
        /// Determines whether a column label and row number collectively represent a valid Excel 
        /// worksheet cell.
        /// </summary>
        /// <remarks>
        /// This method is valid for Excel 2007, Excel 2010 and Excel 2013.
        /// </remarks>
        private static bool CellIsInRange(int row, string col)
        {
            return row >= 1 && row <= 1048576 && col.CompareTo("A") >= 0 && CompareColumns(col, "XFD") <= 0;
        }

        /// <summary>
        /// A comparer method for column labels, facilitating comparisons like "BA" > "Z".
        /// </summary>
        private static int CompareColumns(string col1, string col2)
        {
            int len1 = col1.Length;
            int len2 = col2.Length;
            if (len1 == len2) {
                return col1.CompareTo(col2);
            }
            return len1 < len2 ? -1 : 1;
        }

        /// <summary>
        /// Gets a value indicating whether the range notation supplied during instantiation represents 
        /// a valid Excel worksheet range.
        /// </summary>
        public bool IsValid { get; private set; }

        /// <summary>
        /// Returns an Excel.Range object for a specified worksheet.
        /// </summary>
        /// <param name="worksheet">The Excel.Worksheet for which a range is required.</param>
        /// <exception cref="System.InvalidOperationException">The range notation supplied during 
        /// instantiation is not valid.</exception>
        /// <returns>An Excel.Range implementation.</returns>
        public Excel.Range GetRange(Excel.Worksheet worksheet)
        {
            if (!IsValid) {
                throw new InvalidOperationException("Range notation is not valid.");
            }
            return worksheet.Range[m_range];
        }
    }
}