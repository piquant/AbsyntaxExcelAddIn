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
using AbsyntaxExcelAddIn.Core;
using MI2.Events;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn
{
    /// <summary>
    /// Represents the Absyntax Excel add-in.
    /// </summary>
    public partial class ThisAddIn : IWorksheetProvider
    {
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        /// <summary>
        /// Handler for the add-in's Startup event.
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ReadRules();
            Excel.Application app = Application;
            app.WorkbookDeactivate += Application_WorkbookDeactivate;
            app.WorkbookActivate += Application_WorkbookActivate;
            app.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            SelfDisposingBackgroundWorker.RunWorkerAsync((s, ea) => Start());
        }

        /// <summary>
        /// Handler for the add-in's Shutdown event.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            PerformRuntimeAdapterAction(a => {
                a.ProcessAvailabilityChanged -= RuntimeAdapter_ProcessAvailabilityChanged;
                a.ServiceAvailabilityChanged -= RuntimeAdapter_ServiceAvailabilityChanged;
                a.Dispose();
            });
        }

        /// <summary>
        /// Starts the runtime adapter and monitors its availability.
        /// </summary>
        private void Start()
        {
            Guid? clientId = UsesFullLicence ? null : GetClientId();
            m_runtimeAdapter = new ApplicationRuntimeAdapter(Application, clientId);
            m_serviceAvailable = m_runtimeAdapter.ServiceAvailable;
            m_runtimeAdapter.ProcessAvailabilityChanged += RuntimeAdapter_ProcessAvailabilityChanged;
            m_runtimeAdapter.ServiceAvailabilityChanged += RuntimeAdapter_ServiceAvailabilityChanged;
            Excel.Workbook w = ActiveWorkbook;
            if (w != null) {
                m_runtimeAdapter.SetActiveWorkbook(w.FullName);
            }
            RationaliseRuleIds();
        }

        /// <summary>
        /// Handler for ApplicationRuntimeAdapter.ProcessAvailabilityChanged events.
        /// </summary>
        private void RuntimeAdapter_ProcessAvailabilityChanged(object sender, EventArgs<bool> e)
        {
            m_processAvailable = e.Data;
            OnHostStatusChanged(HostStatusChanged);
        }

        /// <summary>
        /// Handler for ApplicationRuntimeAdapter.ServiceAvailabilityChanged events.
        /// </summary>
        private void RuntimeAdapter_ServiceAvailabilityChanged(object sender, EventArgs<bool> e)
        {
            m_serviceAvailable = e.Data;
            UpdateExecutionState();
            OnHostStatusChanged(HostStatusChanged);
        }

        /// <summary>
        /// The event that is raised whenever the Absyntax runtime process availability or the hosted
        /// service availability changes.
        /// </summary>
        public event EventHandler HostStatusChanged;

        /// <summary>
        /// Invokes the supplied delegate.
        /// </summary>
        private void OnHostStatusChanged(EventHandler handler)
        {
            if (handler != null) {
                handler(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Responsible for handling Absyntax licence details.
        /// </summary>
        private LicenceManager m_licenceManager = new LicenceManager();

        /// <summary>
        /// Returns a nullable Guid representing the id of the client in respect of a third-party execution
        /// licence.
        /// </summary>
        public Guid? GetClientId()
        {
            return m_licenceManager.GetClientId();
        }

        /// <summary>
        /// Sets a nullable Guid representing the id of the client in respect of a third-party execution
        /// licence.
        /// </summary>
        private void SetClientId(Guid? guid)
        {
            m_licenceManager.SetClientId(guid);
        }

        /// <summary>
        /// Gets a value indicating whether a full licence for the Absyntax framework is to be used by this
        /// add-in.
        /// </summary>
        /// <remarks>
        /// If a full licence is available then no client identifier is required in order for the framework
        /// to execute projects on behalf of this add-in.  If a third-party execution licence is the only
        /// available licence then a valid client identifier must be specified.
        /// </remarks>
        public bool UsesFullLicence
        {
            get { return m_licenceManager.UsesFullLicence; }
            private set { m_licenceManager.UsesFullLicence = value; }
        }

        /// <summary>
        /// Updates the cached Absyntax licence details and restarts the runtime if necessary so as to take 
        /// account of these changes.
        /// </summary>
        /// <remarks>
        /// A null client identifier will force the runtime to look for a full licence.
        /// </remarks>
        /// <param name="usesFullLicence">Indicates whether a full licence for the Absyntax framework is to 
        /// be used by this add-in.</param>
        /// <param name="clientId">A nullable Guid representing the id of the client in respect of a 
        /// third-party execution licence.</param>
        public void ChangeLicenceDetails(bool usesFullLicence, Guid? clientId)
        {
            bool d1 = UsesFullLicence != usesFullLicence;
            bool d2 = GetClientId() != clientId;
            if (d1) {
                UsesFullLicence = usesFullLicence;
            }
            if (d2) {
                SetClientId(clientId);
            }
            /* Only restart the runtime adapter if either the UsesFullLicence flag has changed or
             * UsesFullLicence is false and the client identifier GUID has changed.
             * */
            if (d1 || (!usesFullLicence && d2)) {
                Guid? g = usesFullLicence ? null : clientId;
                PerformRuntimeAdapterAction(a => a.Restart(g));
            }
        }

        /// <summary>
        /// An intermediary object that handles all communications between this add-in and the runtime.
        /// </summary>
        private ApplicationRuntimeAdapter m_runtimeAdapter;

        /// <summary>
        /// Handler for events raised when an Excel.Application.Workbook is deactivated.
        /// </summary>
        private void Application_WorkbookDeactivate(Excel.Workbook workbook)
        {
            PerformRuntimeAdapterAction(a => a.DeactivateWorkbook(workbook.FullName));
            HasActiveWorkbook = false;
            UpdateExecutionState();
        }

        /// <summary>
        /// Handler for events raised when an Excel.Application.Workbook is activated.
        /// </summary>
        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            HasActiveWorkbook = true;
            PerformRuntimeAdapterAction(a => a.SetActiveWorkbook(workbook.FullName));
            ReadRules();
        }

        /// <summary>
        /// Handler for events raised immediately before an Excel.Application.Workbook is saved.
        /// </summary>
        private void Application_WorkbookBeforeSave(Excel.Workbook workbook, bool SaveAsUI, ref bool Cancel)
        {
            WriteRules();
        }

        /// <summary>
        /// Execution mode for the active workbook.
        /// </summary>
        private ExecutionMode m_mode;

        /// <summary>
        /// Gets or sets the execution mode for the active workbook.
        /// </summary>
        public ExecutionMode Mode
        {
            get { return m_mode; }
            set { m_mode = value; }
        }

        /// <summary>
        /// Project invocation rules for the active workbook.
        /// </summary>
        private ProjectInvocationRule[] m_rules;

        /// <summary>
        /// Gets or sets the collection of project invocation rules for the active workbook.
        /// </summary>
        public ProjectInvocationRule[] Rules
        {
            get { return m_rules; }
            set {
                m_rules = value;
                UpdateExecutionState();
            }
        }

        /// <summary>
        /// Loads project invocation rules for the active workbook.
        /// </summary>
        private void ReadRules()
        {
            Excel.Workbook workbook = ActiveWorkbook;
            Excel.Worksheet ws = GetRulesWorksheet(workbook);
            if (ws == null) {
                Mode = ExecutionMode.Synchronous;
                Rules = new ProjectInvocationRule[0];
            }
            else {
                var prm = new PersistedRuleManager();
                prm.Load(this, ws, out m_mode, out m_rules);
                RationaliseRuleIds();
                UpdateExecutionState();
            }
        }

        /// <summary>
        /// Rationalises the unique identifiers assigned to the currently loaded project invocation rules.
        /// </summary>
        /// <remarks>
        /// Whan created, each project invocation rule is assigned a number that is unique within the set 
        /// of rules associated with a workbook.  Over times, old rules may be deleted and new ones created.
        /// To prevent rule identifiers from growing unchecked, they are rationalised.
        /// </remarks>
        private void RationaliseRuleIds()
        {
            PerformRuntimeAdapterAction(a => a.RationaliseIds(m_rules));
        }

        /// <summary>
        /// Flags whether the underlying host process is open.  The host process supports the host service
        /// that facilitates project execution.
        /// </summary>
        private bool m_processAvailable = false;

        /// <summary>
        /// Gets a value indicating whether the underlying host process is open.  A process that fails to
        /// open may indicate a licence problem.
        /// </summary>
        public bool ProcessAvailable
        {
            get { return m_processAvailable; }
        }

        /// <summary>
        /// Flags whether the underlying host runtime service is available.
        /// </summary>
        private bool m_serviceAvailable = false;

        /// <summary>
        /// Gets a value indicating whether the underlying host runtime service is available.
        /// </summary>
        public bool ServiceAvailable
        {
            get { return m_serviceAvailable; }
        }

        private bool m_hasActiveWorkbook = true; // Excel starts with an active workbook

        /// <summary>
        /// Gets a value indicating whether there is a workbook available that is the target of this add-in.
        /// </summary>
        /// <remarks>
        /// Excel's Application object has a property named ActiveWorkbook.  This can be non-null even when the
        /// active workbook has been deactivated.  There are no events that can be used to determine when a
        /// workbook has been closed.
        /// </remarks>
        public bool HasActiveWorkbook
        {
            get { return m_hasActiveWorkbook; }
            private set {
                if (m_hasActiveWorkbook != value) {
                    m_hasActiveWorkbook = value;
                    OnStateChanged();
                }
            }
        }

        /// <summary>
        /// Updates this add-in's execution state property.
        /// </summary>
        private void UpdateExecutionState()
        {
            ProjectInvocationRule[] rules = Rules;
            ExecutionState =
                HasActiveWorkbook && 
                m_serviceAvailable && 
                rules != null && 
                rules.Any(r => r.Enabled && r.IsValid) 
                ? AddInExecutionState.CanExecute : AddInExecutionState.CannotExecute;
        }

        /// <summary>
        /// Used to keep track of this add-in's execution state.
        /// </summary>
        private AddInExecutionState m_executionState = AddInExecutionState.CannotExecute;

        /// <summary>
        /// Gets a value indicating the execution state of this add-in.
        /// </summary>
        public AddInExecutionState ExecutionState
        {
            get { return m_executionState; }
            private set {
                if (m_executionState != value) {
                    m_executionState = value;
                    OnStateChanged();
                }
            }
        }

        /// <summary>
        /// The event that is raised whenever this add-in's availability or execution state changes.
        /// </summary>
        public event EventHandler StateChanged;

        /// <summary>
        /// Raises the StateChanged event.
        /// </summary>
        private void OnStateChanged()
        {
            var handler = StateChanged;
            if (handler != null) {
                handler(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// A string used to name the very hidden worksheet that stores project invocation rules for all
        /// workbooks.
        /// </summary>
        private static readonly string s_rulesSheetName = "951BFFBBC625462FA2C0356EE4E5395";

        /// <summary>
        /// Returns the worksheet that stores project invocation rules for a specific workbook.
        /// </summary>
        private static Excel.Worksheet GetRulesWorksheet(Excel.Workbook workbook)
        {
            return workbook.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(s => s.Name == s_rulesSheetName);
        }

        /// <summary>
        /// Stores the active workbook's current project invocation rules within the workbook.
        /// </summary>
        public void WriteRules()
        {
            ProjectInvocationRule[] rules = Rules;
            Excel.Workbook wb = ActiveWorkbook;
            Excel.Worksheet ws = GetRulesWorksheet(wb);
            if (ws == null) {
                if (rules == null || (!rules.Any() && Mode == ExecutionMode.Synchronous)) return;
                ws = (Excel.Worksheet)wb.Worksheets.Add();
                ws.Name = s_rulesSheetName;
                ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            }
            var prm = new PersistedRuleManager();
            prm.Save(ws, Mode, rules);
        }

        /// <summary>
        /// Gets the active workbook.
        /// </summary>
        private Excel.Workbook ActiveWorkbook
        {
            get { return Application.ActiveWorkbook; }
        }

        /// <summary>
        /// Returns an array of worksheets belonging to the active workbook.
        /// </summary>
        public Excel.Worksheet[] GetWorksheets()
        {
            Excel.Workbook wb = ActiveWorkbook;
            return wb.Worksheets.OfType<Excel.Worksheet>().Where(s => s.Visible != Excel.XlSheetVisibility.xlSheetVeryHidden).ToArray();
        }

        /// <summary>
        /// Invokes all valid, enabled project invocation rules for the active workbook.
        /// </summary>
        /// <exception cref="System.InvalidOperationException">The add-in is not in a state that allows 
        /// it to invoke projects.</exception>
        public void Run()
        {
            UpdateRuleValidities();
            UpdateExecutionState();
            if (ExecutionState != AddInExecutionState.CanExecute) return;
            ExecutionState = AddInExecutionState.Executing;
            try {
                PerformRuntimeAdapterAction(a => {
                    a.Run(Mode, Rules);
                    WriteRules();
                });
            }
            finally {
                ExecutionState = AddInExecutionState.CanExecute;
            }
        }

        /// <summary>
        /// Forces each of the active workbook's project invocation rules to update its validity.
        /// </summary>
        private void UpdateRuleValidities()
        {
            ProjectInvocationRule[] rules = Rules;
            if (rules == null) return;
            foreach (ProjectInvocationRule rule in rules) {
                rule.UpdateValidity();
            }
        }

        /// <summary>
        /// Invokes the supplied Action if the application runtime adapter has been instantiated.
        /// </summary>
        private void PerformRuntimeAdapterAction(Action<ApplicationRuntimeAdapter> action)
        {
            ApplicationRuntimeAdapter a = m_runtimeAdapter;
            if (a != null) {
                action(a);
            }
        }
    }
}