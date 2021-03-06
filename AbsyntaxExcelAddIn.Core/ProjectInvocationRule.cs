﻿/* Copyright © 2013-2014 Managing Infrastructure Information Ltd
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AbsyntaxExcelAddIn.Core.Converters;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Encapsulates all details necessary to invoke an Absyntax project.
    /// </summary>
    /// <remarks>
    /// This class also acts a UI view-model for list items in the project configuration dialogue.
    /// </remarks>
    public sealed class ProjectInvocationRule : NotifyPropertyChangedBase, IProjectInvocationRule
    {
        static ProjectInvocationRule()
        {
            var orderings = Enum.GetValues(typeof(RangeOrdering)).Cast<RangeOrdering>();
            var roc = new RangeOrderingConverter();
            s_availableRangeOrderingNames = orderings.Select(o => (string)roc.Convert(o, null, null, null));

            var timeUnits = Enum.GetValues(typeof(TimeUnit)).Cast<TimeUnit>();
            var tuc = new TimeUnitConverter();
            s_availableTimeUnitNames = timeUnits.Select(o => (string)tuc.Convert(o, null, null, null));
        }

        private static readonly IEnumerable<string> s_availableRangeOrderingNames;

        private static readonly IEnumerable<string> s_availableTimeUnitNames;

        private ProjectInvocationRule()
        { }

        /// <summary>
        /// Initialises a new ProjectInvocationRule using persisted field values.
        /// </summary>
        /// <param name="wsProvider">An IWorksheetProvider implementation.</param>
        /// <param name="nrProvider">An INamedRangeProvider implementation.</param>
        /// <param name="reader">An IDataReader capable of supplying the requisite field values on demand.</param>
        /// <exception cref="System.ArgumentNullException">Either the IWorksheetProvider or the INamedRangeProvider 
        /// is null.</exception>
        /// <exception cref="System.NullReferenceException">The IDataReader is null.</exception>
        public ProjectInvocationRule(IWorksheetProvider wsProvider, INamedRangeProvider nrProvider, IDataReader reader)
        {
            SetProviders(wsProvider, nrProvider);
            Id = reader.Read<int>();
            UsesInput = reader.Read<bool>();
            InputSheetKey = reader.Read<string>();
            InputCellRange = reader.Read<string>();
            InputRangeOrder = reader.Read<RangeOrdering>();
            TimeLimit = reader.Read<int>();
            Unit = reader.Read<TimeUnit>();
            UsesOutput = reader.Read<bool>();
            OutputSheetKey = reader.Read<string>();
            OutputCellRange = reader.Read<string>();
            OutputRangeOrder = reader.Read<RangeOrdering>();
            ProjectPath = reader.Read<string>();
            ReloadProjectBeforeExecuting = reader.Read<bool>();
            Enabled = reader.Read<bool>();
            LastExecutionResult = reader.Read<ExecutionResult>();
            UpdateInputSheetRangeNames();
            UpdateOutputSheetRangeNames();
        }

        /// <summary>
        /// Initialises a new, empty ProjectInvocationRule instance.
        /// </summary>
        /// <param name="wsProvider">An IWorksheetProvider implementation.</param>
        /// <param name="nrProvider">An INamedRangeProvider implementation.</param>
        /// <param name="id">A number that identifies the ProjectInvocationRule in a set of such rules.</param>
        /// <exception cref="System.ArgumentNullException">Either the IWorksheetProvider or the INamedRangeProvider 
        /// is null.</exception>
        public ProjectInvocationRule(IWorksheetProvider wsProvider, INamedRangeProvider nrProvider, int id)
        {
            SetProviders(wsProvider, nrProvider);
            Id = id;
            m_inputSheetKey = m_outputSheetKey = GetFirstSheetKey();
            UpdateInputSheetRangeNames();
            UpdateOutputSheetRangeNames();
        }

        /// <summary>
        /// Sets this ProjectInvocationRule's IWorksheetProvider and INamedRangeProvider references, throwing an 
        /// ArgumentNullException if either is null.
        /// </summary>
        private void SetProviders(IWorksheetProvider wsProvider, INamedRangeProvider nrProvider)
        {
            if (wsProvider == null) {
                throw new ArgumentNullException("wsProvider");
            }
            if (nrProvider == null) {
                throw new ArgumentNullException("nrProvider");
            }
            m_wsProvider = wsProvider;
            m_nrProvider = nrProvider;
        }

        /// <summary>
        /// Gets the key of the first worksheet returned by the encapsulated IWorksheetProvider.
        /// </summary>
        private string GetFirstSheetKey()
        {
            Excel.Worksheet worksheet = GetFirstSheet();
            return GetSheetKey(worksheet);
        }

        /// <summary>
        /// An IWorksheetProvider implementation responsible for supplying the available worksheets on
        /// demand.
        /// </summary>
        private IWorksheetProvider m_wsProvider;

        /// <summary>
        /// An INamedRangeProvider implementation responsible for providing services in respect of Excel 
        /// range names.
        /// </summary>
        private INamedRangeProvider m_nrProvider;

        private int m_id;

        /// <summary>
        /// Gets a number that identifies this ProjectInvocationRule in a set of such rules.
        /// </summary>
        public int Id
        {
            get { return m_id; }
            internal set { m_id = value; }
        }

        private bool m_isSelected;

        /// <summary>
        /// Gets or sets a value indicating whether this ProjectInvocationRule is selected in a list of 
        /// such items.
        /// </summary>
        public bool IsSelected
        {
            get { return m_isSelected; }
            set { SetProperty(ref m_isSelected, value, () => IsSelected); }
        }

        public bool m_reloadProjectBeforeExecuting;
        
        /// <summary>
        /// Gets or sets a value indicating whether the represented Absyntax project should be reloaded 
        /// prior to each invocation.
        /// </summary>
        /// <remarks>
        /// For performance reasons it is advantageous to leave a project loaded in memory between 
        /// invocations because project-loading is a relatively slow process.  However, once a project 
        /// is loaded then any changes made to the source project file can only be realised if the project 
        /// is reloaded.  Setting this property to true is particularly useful if you are in the process
        /// of making changes to a project at the same time as using it with this add-in.
        /// </remarks>
        public bool ReloadProjectBeforeExecuting
        {
            get { return m_reloadProjectBeforeExecuting; }
            set { SetProperty(ref m_reloadProjectBeforeExecuting, value, () => ReloadProjectBeforeExecuting); }
        }

        private bool m_enabled = true;

        /// <summary>
        /// Gets or sets a value indicating whether this ProjectInvocationRule is enabled.
        /// </summary>
        /// <remarks>
        /// Disabled rules are not invoked.
        /// </remarks>
        public bool Enabled
        {
            get { return m_enabled; }
            set { SetProperty(ref m_enabled, value, () => Enabled); }
        }

        private bool m_usesInput;

        /// <summary>
        /// Gets or sets a value indicating whether data will be obtained from a worksheet and passed to 
        /// the represented Absyntax project before each invocation.
        /// </summary>
        /// <remarks>
        /// Some Absyntax projects may have an entry-point that requires data, others may not.  Users are 
        /// reasonably expected to know these details.
        /// </remarks>
        public bool UsesInput
        {
            get { return m_usesInput; }
            set {
                if (SetProperty(ref m_usesInput, value, () => UsesInput)) {
                    UpdateValidityInternal();
                }
            }
        }

        private string m_inputSheetKey;

        /// <summary>
        /// Gets the unique key of the worksheet that will be used to provide input data to the represented 
        /// Absyntax project when UsesInput is set to true.
        /// </summary>
        public string InputSheetKey
        {
            get { return m_inputSheetKey; }
            private set { m_inputSheetKey = value; }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet that will be used to provide input data to the represented 
        /// Absyntax project when UsesInput is set to true.
        /// </summary>
        public string InputSheetName
        {
            get { return GetSheetName(m_inputSheetKey); }
            set {
                Excel.Worksheet ws = GetSheetByName(value);
                m_inputSheetKey = GetSheetKey(ws);
                EnsureCompatibleCellRange(m_inputSheetKey, InputCellRange, () => InputCellRange = String.Empty);
                OnPropertyChanged(() => InputSheetName);
                UpdateValidityInternal();
            }
        }

        /// <summary>
        /// Ensures that a cell range is compatible with a currently selected sheet name.
        /// </summary>
        private void EnsureCompatibleCellRange(string sheetKey, string cellRange, Action action)
        {
            Excel.Worksheet ws = m_nrProvider.IdentifyWorksheet(cellRange);
            if (ws == null) return;
            string otherKey = GetSheetKey(ws);
            if (sheetKey != otherKey) {
                action();
            }
        }

        /// <summary>
        /// Updates the list of available input sheet range names for the currently selected input worksheet 
        /// name.
        /// </summary>
        private void UpdateInputSheetRangeNames()
        {
            string[] rangeNames = m_nrProvider.GetRangeNames();
            InputSheetRangeNames = rangeNames;
        }

        private string m_inputCellRange = "A1:B2";

        /// <summary>
        /// Gets or sets a notation defining a range of cells that will be used to obtain input data to be 
        /// passed to the represented Absyntax project before each invocation when UsesInput is set to true.
        /// </summary>
        /// <remarks>
        /// This property supports range names.  Because a named range may consist of multiple areas, this 
        /// means that it is possible to read input data from non-contiguous cells.  In such cases, data is 
        /// read from areas in the order they are defined within the named range.
        /// </remarks>
        public string InputCellRange
        {
            get { return m_inputCellRange; }
            set {
                if (SetProperty(ref m_inputCellRange, value, () => InputCellRange)) {
                    EnsureCompatibleInputSheet();
                    UpdateValidityInternal();
                }
            }
        }

        private RangeOrdering m_inputRangeOrder;

        /// <summary>
        /// Gets or sets a value that determines the order in which the input data range of cell values are to
        /// be offered to the represented Absyntax project before each invocation when UsesInput is set to true.
        /// </summary>
        public RangeOrdering InputRangeOrder
        {
            get { return m_inputRangeOrder; }
            set { SetProperty(ref m_inputRangeOrder, value, () => InputRangeOrder); }
        }

        private bool m_usesOutput;

        /// <summary>
        /// Gets or sets a value indicating whether any data output by the represented Absyntax project will 
        /// be written to a worksheet after each invocation.
        /// </summary>
        /// <remarks>
        /// Some Absyntax projects may have an exit-point that passes data, others may not.  Users are at 
        /// liberty to ignore a project's output data if they want to.
        /// </remarks>
        public bool UsesOutput
        {
            get { return m_usesOutput; }
            set {
                if (SetProperty(ref m_usesOutput, value, () => UsesOutput)) {
                    UpdateValidityInternal();
                }
            }
        }

        private string m_outputSheetKey;

        /// <summary>
        /// Gets the unique key of the worksheet that will be used to write output data received from the 
        /// represented Absyntax project when UsesOutput is set to true.
        /// </summary>
        public string OutputSheetKey
        {
            get { return m_outputSheetKey; }
            private set { m_outputSheetKey = value; }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet that will be used to write output data received from the 
        /// represented Absyntax project when UsesOutput is set to true.
        /// </summary>
        public string OutputSheetName
        {
            get { return GetSheetName(m_outputSheetKey); }
            set {
                Excel.Worksheet ws = GetSheetByName(value);
                m_outputSheetKey = GetSheetKey(ws);
                EnsureCompatibleCellRange(m_outputSheetKey, OutputCellRange, () => OutputCellRange = String.Empty);
                OnPropertyChanged(() => OutputSheetName);
                UpdateValidityInternal();
            }
        }

        private string m_outputCellRange = "C1";

        /// <summary>
        /// Gets or sets a notation defining a range of cells that will be used to write data received from 
        /// the represented Absyntax project after each invocation when UsesOutput is set to true.
        /// </summary>
        /// <remarks>
        /// Note that a project's output is not confined to the specified range.  For example, if the range
        /// defines an area of two columns and five rows (i.e. ten cells) and the project outputs an array of
        /// 11 values, all 11 values will be written.  This is achieved by extending the range in a direction 
        /// determined by the value of the OutputRangeOrder property.
        /// <para />
        /// In fact, the output cell range needs only to be defined in terms of 1 x n columns or n x 1 rows.
        /// For example, an output range defined as "C6:G6" with an output range ordering equal to 
        /// RangeOrdering.ByRow will result in the first five output values being written to cells C6 through 
        /// G6, the next five being written to cells C7 through G7 and so on.  Similarly, an output range 
        /// defined as "C6:C10" with an output range ordering equal to RangeOrdering.ByColumn will result in 
        /// the first five output values being written to cells C6 through C10, the next five being written to 
        /// cells D6 through D10 and so on.
        /// <para />
        /// This property also supports range names.  Because a named range may consist of multiple areas, this 
        /// means that it is possible to write output data to non-contiguous cells.  In such cases, data is 
        /// written to areas in the order they are defined within the named range.  Only the last area may be 
        /// extended as described above.  All other areas are filled up in turn.
        /// </remarks>
        public string OutputCellRange
        {
            get { return m_outputCellRange; }
            set {
                if (SetProperty(ref m_outputCellRange, value, () => OutputCellRange)) {
                    EnsureCompatibleOutputSheet();
                    UpdateValidityInternal();
                }
            }
        }

        /// <summary>
        /// Updates the list of available output sheet range names for the currently selected output worksheet 
        /// name.
        /// </summary>
        private void UpdateOutputSheetRangeNames()
        {
            string[] rangeNames = m_nrProvider.GetRangeNames();
            OutputSheetRangeNames = rangeNames;
        }

        private RangeOrdering m_outputRangeOrder;

        /// <summary>
        /// Gets or sets a value that determines the order in which cells are to be written using output data 
        /// received from the represented Absyntax project are to be written after each invocation when 
        /// UsesOutput is set to true.
        /// </summary>
        public RangeOrdering OutputRangeOrder
        {
            get { return m_outputRangeOrder; }
            set { SetProperty(ref m_outputRangeOrder, value, () => OutputRangeOrder); }
        }

        private int m_timeLimit = 10;

        /// <summary>
        /// Gets or sets a value which, when combined with the Unit property value, determines the amount of 
        /// time that Absyntax will allow for a project invocation to complete before terminating an 
        /// invocation.
        /// </summary>
        public int TimeLimit
        {
            get { return m_timeLimit; }
            set {
                if (value > 0) {
                    m_timeLimit = value;
                }
                OnPropertyChanged(() => TimeLimit);
            }
        }

        private TimeUnit m_timeUnit;

        /// <summary>
        /// Gets or sets a value which, when combined with the TimeLimit property value, determines the amount 
        /// of time that Absyntax will allow for a project invocation to complete before terminating an 
        /// invocation.
        /// </summary>
        public TimeUnit Unit
        {
            get { return m_timeUnit; }
            set { SetProperty(ref m_timeUnit, value, () => Unit); }
        }

        private string m_projectPath;

        /// <summary>
        /// Gets or sets the full path of the file containing the serialised form of the Absyntax project to
        /// be invoked.
        /// </summary>
        public string ProjectPath
        {
            get { return m_projectPath; }
            set {
                if (SetProperty(ref m_projectPath, value, () => ProjectPath)) {
                    UpdateValidityInternal();
                }
            }
        }

        /// <summary>
        /// Returns the first available worksheet.
        /// </summary>
        private Excel.Worksheet GetFirstSheet()
        {
            return m_wsProvider.GetWorksheets().FirstOrDefault();
        }

        /// <summary>
        /// Returns the worksheet whose name matches the supplied value.
        /// </summary>
        private Excel.Worksheet GetSheetByName(string name)
        {
            return m_wsProvider.GetWorksheets().FirstOrDefault(w => w.Name == name);
        }

        /// <summary>
        /// Returns the worksheet whose unique key matches the supplied value.
        /// </summary>
        private Excel.Worksheet GetSheetByKey(string key)
        {
            var wi = new WorksheetIdentifier();
            return wi.GetWorksheet(m_wsProvider, key);
        }

        /// <summary>
        /// Returns the name of the worksheet whose unique key is supplied.
        /// </summary>
        private string GetSheetName(string key)
        {
            Excel.Worksheet ws = GetSheetByKey(key);
            return ws == null ? null : ws.Name;
        }

        /// <summary>
        /// Returns the unique key of the supplied worksheet.  If the worksheet does not have one then one is
        /// created and assigned.
        /// </summary>
        private string GetSheetKey(Excel.Worksheet worksheet)
        {
            return worksheet == null ? null : new WorksheetIdentifier().GetKey(worksheet);
        }

        /// <summary>
        /// Gets a collection of the names of the available worksheets.
        /// </summary>
        public IEnumerable<string> AvailableSheetNames
        {
            get { return m_wsProvider.GetWorksheets().Select(w => w.Name); }
        }

        /// <summary>
        /// Gets a collection of the available range names for the currently selected input worksheet name.
        /// </summary>
        public IEnumerable<string> InputSheetRangeNames { get; private set; }

        /// <summary>
        /// Gets a collection of the available range names for the currently selected output worksheet name.
        /// </summary>
        public IEnumerable<string> OutputSheetRangeNames { get; private set; }

        /// <summary>
        /// Gets a collection of the names of the range-ordering options.
        /// </summary>
        public IEnumerable<string> AvailableRangeOrderingNames
        {
            get { return s_availableRangeOrderingNames; }
        }

        /// <summary>
        /// Gets a collection of the names of the time units.
        /// </summary>
        public IEnumerable<string> AvailableTimeUnits
        {
            get { return s_availableTimeUnitNames; }
        }

        /// <summary>
        /// Updates the IsValid property value based on the state of this ProjectInvocationRule.
        /// </summary>
        private void UpdateValidityInternal()
        {
            UpdateValidity(false);
        }

        /// <summary>
        /// Ensures that selected input and output sheets are compatible with their respective named cell 
        /// ranges and updates the IsValid property value based on the state of this ProjectInvocationRule.
        /// </summary>
        public void UpdateValidity()
        {
            UpdateValidity(true);
        }

        /// <summary>
        /// Ensures that selected input and output sheets are compatible with their respective named cell 
        /// ranges and updates the IsValid property value based on the state of this ProjectInvocationRule.
        /// </summary>
        /// <param name="ensureCompatibleSheets">Determines whether operations should be performed to ensure
        /// that the input and output sheet names are compatible with the current input and output cell ranges.</param>
        private void UpdateValidity(bool ensureCompatibleSheets)
        {
            if (ensureCompatibleSheets) {
                EnsureCompatibleInputSheet();
                EnsureCompatibleOutputSheet();
            }
            IsValid = DetermineValidity();
        }

        /// <summary>
        /// Ensures that the selected input sheet is compatible with the input cell range.
        /// </summary>
        /// <remarks>
        /// If the input cell range references a range name, this method ensures that the selected input
        /// sheet name corresponds with the name of the worksheet that is referenced by said name.  Excel 
        /// allows the user to change the areas of a named range.
        /// </remarks>
        private void EnsureCompatibleInputSheet()
        {
            EnsureCompatibleSheet(InputCellRange, ws => InputSheetName = ws.Name);
        }

        /// <summary>
        /// Ensures that the selected output sheet is compatible with the output cell range.
        /// </summary>
        /// <remarks>
        /// If the output cell range references a range name, this method ensures that the selected output
        /// sheet name corresponds with the name of the worksheet that is referenced by said name.  Excel 
        /// allows the user to change the areas of a named range.
        /// </remarks>
        private void EnsureCompatibleOutputSheet()
        {
            EnsureCompatibleSheet(OutputCellRange, ws => OutputSheetName = ws.Name);
        }

        /// <summary>
        /// Invokes an Action if a range name cannot be associated with an existing worksheet.
        /// </summary>
        private void EnsureCompatibleSheet(string cellRange, Action<Excel.Worksheet> action)
        {
            Excel.Worksheet ws = m_nrProvider.IdentifyWorksheet(cellRange);
            if (ws != null) {
                action(ws);
            }
        }

        /// <summary>
        /// Returns a value indicating whether this ProjectInvocationRule is in a valid state.
        /// </summary>
        private bool DetermineValidity()
        {
            if (UsesInput) {
                Excel.Worksheet ws = GetSheetByKey(m_inputSheetKey);
                if (ws == null) {
                    return false;
                }
                var v = new CellRangeValidator(InputCellRange, m_nrProvider);
                if (!v.IsValid) {
                    return false;
                }
            }
            if (UsesOutput) {
                Excel.Worksheet ws = GetSheetByKey(m_outputSheetKey);
                if (ws == null) {
                    return false;
                }
                var v = new CellRangeValidator(OutputCellRange, m_nrProvider);
                if (!v.IsValid) {
                    return false;
                }
            }
            if (!PathIsValid(ProjectPath)) {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Determines whether a file path represents an existing file.
        /// </summary>
        private bool PathIsValid(string path)
        {
            if (path == null) return false;
            try {
                var fi = new FileInfo(path);
                return fi.Exists;
            }
            catch { }
            return false;
        }

        private bool m_isValid;

        /// <summary>
        /// Gets a value indicating whether this ProjectInvocationRule is in a state that permits it to be 
        /// invoked.
        /// </summary>
        public bool IsValid
        {
            get { return m_isValid; }
            private set { SetProperty(ref m_isValid, value, () => IsValid); }
        }

        private ExecutionResult m_lastExecutionResult;
        
        /// <summary>
        /// Gets this ProjectInvocationRule's last execution result.
        /// </summary>
        public ExecutionResult LastExecutionResult
        {
            get { return m_lastExecutionResult; }
            internal set { SetProperty(ref m_lastExecutionResult, value, () => LastExecutionResult); }
        }

        /// <summary>
        /// Gets a value indicating whether this rule is in a state that allows the Absyntax project it 
        /// represents to be executed.
        /// </summary>
        public bool CanExecute
        {
            get { return IsValid && Enabled; }
        }

        /// <summary>
        /// Reads a collection of values from cells in the input range of the selected input data worksheet, 
        /// sequenced in accordance with the input range ordering value.
        /// </summary>
        /// <returns>A collection of cell values.</returns>
        /// <exception cref="System.InvalidOperationException">Either the UsesInput property value is 
        /// false or this ProjectInvocationRule is not in a valid state.</exception>
        public IEnumerable<object> ReadInputData()
        {
            if (!UsesInput || !IsValid) {
                throw new InvalidOperationException();
            }
            var v = new CellRangeValidator(InputCellRange, m_nrProvider);
            Excel.Worksheet ws = GetSheetByKey(m_inputSheetKey);
            Excel.Range r = v.GetRange(ws);
            return Helper.GetRangeValues(r, InputRangeOrder);
        }

        /// <summary>
        /// Attempts to write an object to a range of cells defined by the various output-related state values
        /// of this ProjectInvocationRule.
        /// </summary>
        /// <remarks>
        /// If the object is enumerable, its enumerated values are written.
        /// </remarks>
        /// <exception cref="System.InvalidOperationException">The UsesOutput property value is false.</exception>
        public void WriteOutputData(object data)
        {
            if (!UsesOutput) {
                throw new InvalidOperationException();
            }
            var v = new CellRangeValidator(OutputCellRange, m_nrProvider);
            Excel.Worksheet ws = GetSheetByKey(m_outputSheetKey);
            Excel.Range r = v.GetRange(ws);
            IEnumerable e = data as IEnumerable;
            if (e == null || e is string) {
                ws.Cells[r.Row, r.Column] = data;
            }
            else {
                Helper.SetRangeValues(e.Cast<object>(), r, OutputRangeOrder);
            }
        }

        /// <summary>
        /// Clones this ProjectInvocationRule.
        /// </summary>
        /// <returns>A new ProjectInvocationRule instance representing a clone of this ProjectInvocationRule.</returns>
        public ProjectInvocationRule Clone()
        {
            var rule = new ProjectInvocationRule() {
                m_wsProvider = this.m_wsProvider,
                m_nrProvider = this.m_nrProvider,
                m_id = this.m_id,
                m_usesInput = this.m_usesInput,
                m_inputSheetKey = this.m_inputSheetKey,
                m_inputCellRange = this.m_inputCellRange,
                m_inputRangeOrder = this.m_inputRangeOrder,
                m_timeLimit = this.m_timeLimit,
                m_timeUnit = this.m_timeUnit,
                m_usesOutput = this.m_usesOutput,
                m_outputSheetKey = this.m_outputSheetKey,
                m_outputCellRange = this.m_outputCellRange,
                m_outputRangeOrder = this.m_outputRangeOrder,
                m_projectPath = this.m_projectPath,
                m_reloadProjectBeforeExecuting = this.m_reloadProjectBeforeExecuting,
                m_enabled = this.m_enabled,
                m_lastExecutionResult = this.m_lastExecutionResult
            };
            rule.UpdateInputSheetRangeNames();
            rule.UpdateOutputSheetRangeNames();
            rule.UpdateValidity(true);
            return rule;
        }

        /// <summary>
        /// Serialises this ProjectInvocationRule.
        /// </summary>
        /// <param name="writer">An IDataWriter to which field values are written.</param>
        public void Write(IDataWriter writer)
        {
            writer.Write(Id);
            writer.Write(UsesInput);
            writer.Write(InputSheetKey);
            writer.Write(InputCellRange);
            writer.Write(InputRangeOrder.ToString());
            writer.Write(TimeLimit);
            writer.Write(Unit.ToString());
            writer.Write(UsesOutput);
            writer.Write(OutputSheetKey);
            writer.Write(OutputCellRange);
            writer.Write(OutputRangeOrder.ToString());
            writer.Write(ProjectPath);
            writer.Write(ReloadProjectBeforeExecuting);
            writer.Write(Enabled);
            writer.Write(LastExecutionResult.ToString());
        }
    }
}