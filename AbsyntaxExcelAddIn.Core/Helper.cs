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
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using MI2.TypeConversion;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A static collection of miscellaneous helper methods.
    /// </summary>
    public static class Helper
    {
        /// <summary>
        /// Ensures that a WPF application is started and that required application-level resources are 
        /// made available.
        /// </summary>
        public static void EnsureApplicationResources()
        {
            if (Application.Current == null) {
                // Create the WPF Application object...
                new Application();

                // ... then merge in application resources.
                Application.Current.Resources.MergedDictionaries.Add(
                    Application.LoadComponent(
                        new Uri("AbsyntaxExcelAddIn.Core;component/Themes/ExpressionDark.xaml", UriKind.Relative)) as ResourceDictionary);
            }
        }

        /// <summary>
        /// Returns a cell descriptor of the form "A1" for a pair of 1-based column and row indices.
        /// </summary>
        /// <remarks>
        /// For example:
        /// GetCell(1, 1) returns "A1".
        /// GetCell(28, 10) returns "AB10".
        /// </remarks>
        public static string GetCell(int colIndex, int row)
        {
            if (colIndex < 1 || row < 1) {
                throw new ArgumentOutOfRangeException();
            }
            string col = GetCol(colIndex);
            return GetCell(col, row);
        }

        /// <summary>
        /// ASCII code for 'A'.
        /// </summary>
        private static readonly int s_ascA = (int)'A';

        /// <summary>
        /// Returns the name of a column whose 1-based index is supplied.
        /// </summary>
        /// <remarks>
        /// For example:
        /// GetCell(1) returns "A".
        /// GetCell(28) returns "AB".
        /// </remarks>
        private static string GetCol(int index)
        {
            var list = new List<char>();
            return GetCol(index, list);
        }

        private static string GetCol(int num, IList<char> list)
        {
            int a = ((num - 1) % 26);
            list.Add((char)(a + s_ascA));
            num = (num - 1) / 26;
            return num > 0 ? GetCol(num, list) : new string(list.Reverse().ToArray());
        }

        /// <summary>
        /// Returns the concatenated form of a column name and a row index.
        /// </summary>
        public static string GetCell(string col, int row)
        {
            return String.Format("{0}{1}", col, row);
        }

        /// <summary>
        /// Writes an object to a worksheet cell identified by 1-based row and column indices.
        /// </summary>
        /// <param name="ws">The Excel worksheet containing the cell to be written.</param>
        /// <param name="colIndex">The 1-based index of the column containing the cell to be written.</param>
        /// <param name="row">The 1-based index of the row containing the cell to be written.</param>
        /// <param name="value">The value to be written to the cell.</param>
        public static void WriteCell(Excel.Worksheet ws, int colIndex, int row, object value)
        {
            string cell = GetCell(colIndex, row);
            ws.Range[cell].Value2 = value;
        }

        /// <summary>
        /// Reads a value from a worksheet cell.
        /// </summary>
        /// <typeparam name="T">The type of value to be returned.</typeparam>
        /// <param name="ws">The Excel worksheet containing the cell to be read.</param>
        /// <param name="colIndex">The 1-based index of the column containing the cell to be written.</param>
        /// <param name="row">The 1-based index of the row containing the cell to be written.</param>
        /// <returns>The cell value as an object of type T.</returns>
        public static T ReadCell<T>(Excel.Worksheet ws, int colIndex, int row)
        {
            T value = default(T);
            string cell = GetCell(colIndex, row);
            object cv = ws.Range[cell].Value2;
            if (cv != null) {
                bool failed;
                value = (T)ConverterFactory.Convert(cv, typeof(T), out failed);
            }
            return value;
        }

        /// <summary>
        /// Creates a 1-based integer that does not exist in a collection of such numbers.
        /// </summary>
        /// <param name="existingIds">A collection of numbers that are not to be replicated.</param>
        public static int CreateId(IEnumerable<int> existingIds)
        {
            int id = 1;
            while (existingIds.Contains(id)) {
                id++;
            }
            return id;
        }

        /// <summary>
        /// Converts a number of time units into a number of milliseconds.
        /// </summary>
        public static int GetMilliseconds(int timeLimit, TimeUnit unit)
        {
            switch (unit) {
                case TimeUnit.Seconds:
                    return timeLimit * 1000;
                case TimeUnit.Minutes:
                    return timeLimit * 60000;
                case TimeUnit.Hours:
                    return timeLimit * 3600000;
                default: // Days
                    return timeLimit * 86400000;
            }
        }

        /// <summary>
        /// Invokes the supplied action on the supplied Dispatcher's thread.
        /// </summary>
        public static void PerformDispatcherAction(Dispatcher dispatcher, Action action)
        {
            if (dispatcher.CheckAccess()) {
                action();
            }
            else {
                dispatcher.BeginInvoke(action, DispatcherPriority.DataBind);
            }
        }

        /// <summary>
        /// Appends one string to another, inserting a couple of new-line characters between them.
        /// </summary>
        /// <param name="text1">The string to which a string paragraph is to be appended.  This string is 
        /// updated in place.</param>
        /// <param name="text2">The paragraph to be appended.</param>
        public static void AddParagraph(ref string text1, string text2)
        {
            text1 += String.Format("{0}{0}{1}", Environment.NewLine, text2);
        }

        /// <summary>
        /// Converts an Excel range into a collection of cell values.
        /// </summary>
        /// <param name="range">The range whose cell values are to be obtained.</param>
        /// <param name="order">Determines whether cell values are presented by row or by column in the 
        /// returned collection.</param>
        /// <returns>A sequence of objects representing the values of the cells in the range.</returns>
        public static IEnumerable<object> GetRangeValues(Excel.Range range, RangeOrdering order)
        {
            if (range.Count == 1) {
                return new object[] { range.Value2 };
            }
            object[,] arr = (object[,])range.Value2;
            if (order == RangeOrdering.ByColumn) {
                arr = SwapRowsAndCols(arr);
            }
            return arr.Cast<object>();
        }

        /// <summary>
        /// Swaps the rows and columns of a two-dimensional array with one-based ranks.
        /// </summary>
        /// <param name="arr">The two-dimensional array whose rows and columns are to be swapped.</param>
        /// <returns>A new two-dimensional array with zero-based ranks.</returns>
        private static object[,] SwapRowsAndCols(object[,] arr)
        {
            int rowCount = arr.GetUpperBound(1);
            int colCount = arr.GetUpperBound(0);
            object[,] t = new object[rowCount, colCount];
            for (int row = 0; row < rowCount; row++) {
                for (int col = 0; col < colCount; col++) {
                    t[row, col] = arr[col + 1, row + 1];
                }
            }
            return t;
        }

        /// <summary>
        /// Writes all of the values of the supplied collection to an Excel range.
        /// </summary>
        /// <param name="e">The IEnumerable whose enumerated values are to be written to the range.</param>
        /// <param name="range">An Excel.Range to be used as the basis for defining a new range, in 
        /// the same worksheet location (i.e. with the same top-left cell) and which is sufficiently 
        /// large to accommodate all enumerated values.</param>
        /// <param name="order">A RangeOrdering that determines the next cell in the range to be 
        /// written to.</param>
        public static void SetRangeValues(IEnumerable<object> e, Excel.Range range, RangeOrdering order)
        {
            int valueCount = e.Count();
            int rowCount, colCount;
            if (order == RangeOrdering.ByRow) {
                colCount = range.Columns.Count;
                rowCount = valueCount / colCount;
                if (valueCount % colCount > 0) {
                    rowCount++;
                }
            }
            else {
                rowCount = range.Rows.Count;
                colCount = valueCount / rowCount;
                if (valueCount % rowCount > 0) {
                    colCount++;
                }
            }
            range = range.Resize[rowCount, colCount];
            object[,] arr = Get2DArray(e, rowCount, colCount, order);
            range.Value2 = arr;
        }

        /// <summary>
        /// Converts a generic IEnumerable into a two-dimensional object array.
        /// </summary>
        /// <param name="e">The generic IEnumerable to be converted.</param>
        /// <param name="rowCount">The number of rows to be supported by the new array.</param>
        /// <param name="colCount">The number of columns to be supported by the new array.</param>
        /// <param name="order">A RangeOrdering that determines the next cell in the range to be 
        /// written to.</param>
        /// <returns>A two-dimensional object array.</returns>
        private static object[,] Get2DArray(IEnumerable<object> e, int rowCount, int colCount, RangeOrdering order)
        {
            object[,] arr = (object[,])Array.CreateInstance(typeof(object), new int[] { rowCount, colCount }, new int[] { 1, 1 });
            int row = 1, col = 1;
            if (order == RangeOrdering.ByRow) {
                foreach (object value in e) {
                    arr[row, col++] = value;
                    if (col > colCount) {
                        col = 1;
                        row++;
                    }
                }
            }
            else {
                foreach (object value in e) {
                    arr[row++, col] = value;
                    if (row > rowCount) {
                        row = 1;
                        col++;
                    }
                }
            }
            return arr;
        }
    }
}