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
using System.Runtime.InteropServices;
using System.Timers;
using MI2.Events;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// This class acts as the buffer between the Absyntax Excel add-in and the Absyntax framework's
    /// isolated runtime adapter.
    /// </summary>
    public sealed class ApplicationRuntimeAdapter : IDisposable
    {
        /// <summary>
        /// Initialises a new ApplicationRuntimeAdapter instance.
        /// </summary>
        /// <param name="application">A reference to the Excel application.</param>
        /// <param name="clientId">A nullable Guid representing the unique identifier of a client of this 
        /// Excel add-in, which must relate to an existing, activated Absyntax third-party execution 
        /// licence.  Set this to null if a full licence is to be used.</param>
        /// <exception cref="System.ArgumentNullException">The application is null.</exception>
        public ApplicationRuntimeAdapter(Excel.Application application, Guid? clientId)
        {
            if (application == null) {
                throw new ArgumentNullException("application");
            }
            m_application = application;
            m_manager = new IsolatedRuntimeManager();
            m_manager.ServiceAvailabilityChanged += Manager_ServiceAvailabilityChanged;
            m_timer = new Timer(60000.0) { Enabled = true };
            m_timer.Elapsed += Timer_Elapsed;
            SelfDisposingBackgroundWorker.RunWorkerAsync((s, ea) => StartRuntime(clientId));
        }

        /// <summary>
        /// A reference to the Excel application.
        /// </summary>
        private Excel.Application m_application;

        /// <summary>
        /// Flags whether a flush request is in progress.
        /// </summary>
        private bool m_flushPending;

        /// <summary>
        /// Handler for the IsolatedRuntimeManager.ServiceAvailabilityChanged event.
        /// </summary>
        /// <remarks>
        /// If the Absyntax host service becomes unavailable, all WorkbookRuntimeAdapters should be 
        /// "flushed" because any identifiers they have for loaded Absyntax projects are now
        /// invalid.
        /// </remarks>
        private void Manager_ServiceAvailabilityChanged(object sender, EventArgs<bool> e)
        {
            if (!e.Data && !m_flushPending) {
                SelfDisposingBackgroundWorker.RunWorkerAsync((s, a) => {
                    m_flushPending = true;
                    try {
                        lock (m_runLock) {
                            FlushAdapters();
                        }
                    }
                    finally {
                        m_flushPending = false;
                    }
                });
            }
        }

        /// <summary>
        /// Causes all registered WorkbookRuntimeAdapters to unload any references to loaded projects.
        /// </summary>
        private void FlushAdapters()
        {
            foreach (WorkbookRuntimeAdapter a in m_workbookAdapters.Values) {
                a.UnloadAll(m_manager);
            }
        }

        /// <summary>
        /// A <see cref="System.Timers.Timer"/> for invoking an operation to dispose of redundant 
        /// WorkbookRuntimeAdapter instances.
        /// </summary>
        private Timer m_timer;

        /// <summary>
        /// Handles the <see cref="System.Timers.Timer"/>.Elapsed event.
        /// </summary>
        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            m_timer.Stop();
            try {
                TrimAdapters();
            }
            finally {
                m_timer.Start();
            }
        }

        /// <summary>
        /// Attempts to close and re-open the Absyntax runtime service.
        /// </summary>
        private void StartRuntime(Guid? clientId)
        {
            lock (m_manager) {
                m_manager.Close();
                m_manager.Open(clientId);
            }
        }

        private IsolatedRuntimeManager m_manager;

        /// <summary>
        /// Restarts the Absyntax Runtime Server.
        /// </summary>
        /// <remarks>
        /// Call this method if (a) an Absyntax third-party execution licence is to be used by this add-in 
        /// (typically because you do not have a full licence), (b) a different third-party licence is to 
        /// be used than the one being used at the moment, or (c) a full licence is to be used when a 
        /// third-party licence was previously being used.
        /// </remarks>
        /// <param name="clientId">A nullable Guid representing the unique identifier of a client of this 
        /// Excel add-in which must relate to an existing, activated Absyntax third-party execution 
        /// licence.  Set this to null if a full licence is to be used.</param>
        public void Restart(Guid? clientId)
        {
            StartRuntime(clientId);
        }

        /// <summary>
        /// The event that is raised whenever the Absyntax runtime process availability changes.
        /// </summary>
        public event EventHandler<EventArgs<bool>> ProcessAvailabilityChanged
        {
            add { m_manager.ProcessAvailabilityChanged += value; }
            remove { m_manager.ProcessAvailabilityChanged -= value; }
        }

        /// <summary>
        /// Gets a value indicating whether the Absyntax runtime process is open.
        /// </summary>
        public bool ProcessAvailable
        {
            get { return m_manager.ProcessAvailable; }
        }

        /// <summary>
        /// The event that is raised whenever the Absyntax runtime process's hosting service availability 
        /// changes.
        /// </summary>
        public event EventHandler<EventArgs<bool>> ServiceAvailabilityChanged
        {
            add { m_manager.ServiceAvailabilityChanged += value; }
            remove { m_manager.ServiceAvailabilityChanged -= value; }
        }

        /// <summary>
        /// Gets a value indicating whether the Absyntax runtime process's hosting service is available.
        /// </summary>
        public bool ServiceAvailable
        {
            get { return m_manager.ServiceAvailable; }
        }

        /// <summary>
        /// Advises this ApplicationRuntimeAdapter of the newly activated workbook's full name.
        /// </summary>
        /// <param name="name">The full name of the activated workbook.</param>
        public void SetActiveWorkbook(string name)
        {
            m_activeWorkbookName = name;
            if (name == null) {
                m_activeAdapter = null;
            }
            else {
                m_activeAdapter = GetAdapter(name);
            }
        }

        /// <summary>
        /// The full name of the active workbook.
        /// </summary>
        /// <remarks>
        /// It is possible that the active workbook's name may change while it is active.  This 
        /// ApplicationRuntimeAdapter only needs to be told of any name change when the workbook is 
        /// deactivated.
        /// </remarks>
        private string m_activeWorkbookName;

        /// <summary>
        /// A reference to the WorkbookRuntimeAdapter associated with the application's active 
        /// workbook.
        /// </summary>
        private WorkbookRuntimeAdapter m_activeAdapter;

        /// <summary>
        /// Advises this ApplicationRuntimeAdapter that the most recently active workbook has been
        /// deactivated.
        /// </summary>
        /// <remarks>
        /// The active workbook's name may change.  Because this ApplicationRuntimeAdapter uses a 
        /// workbook's full name to identify the workbook's associated WorkbookRuntimeAdapter, it must be 
        /// advised of any name change.  Unfortunately there is no Excel interop event for this.
        /// </remarks>
        /// <param name="name">The full name of the deactivated workbook.</param>
        public void DeactivateWorkbook(string name)
        {
            string awn = m_activeWorkbookName;
            if (awn != null && awn != name) {
                // The active workbook's name has been changed
                MapAdapter(awn, name);
            }
            m_activeWorkbookName = null;
            m_activeAdapter = null;
        }

        /// <summary>
        /// Updates the workbook adapter cache by re-keying the adapter associated with a workbook whose 
        /// full name has changed.
        /// </summary>
        private void MapAdapter(string oldName, string newName)
        {
            lock (m_workbookAdapters) {
                WorkbookRuntimeAdapter a;
                if (m_workbookAdapters.TryGetValue(oldName, out a)) {
                    m_workbookAdapters.Remove(oldName);
                    m_workbookAdapters[newName] = a;
                }
            }
        }

        /// <summary>
        /// Invokes the supplied action if the currently active WorkbookRuntimeAdapter is not null.
        /// </summary>
        private void PerformActiveAdapterAction(Action<WorkbookRuntimeAdapter> action)
        {
            WorkbookRuntimeAdapter a = m_activeAdapter;
            if (a != null) {
                action(a);
            }
        }

        /// <summary>
        /// Re-bases the unique identifiers assigned to a set of ProjectInvocationRule instances.
        /// </summary>
        /// <param name="rules">An array of ProjectInvocationRules whose identifiers are to be
        /// rationalised.</param>
        public void RationaliseIds(ProjectInvocationRule[] rules)
        {
            PerformActiveAdapterAction(a => a.RationaliseIds(rules));
        }

        /// <summary>
        /// Used to ensure that adapters are not flushed through service unavailability while a run is 
        /// in progress.
        /// </summary>
        private object m_runLock = new object();

        /// <summary>
        /// Invokes those of the supplied project invocation rules that are valid and enabled.
        /// </summary>
        /// <param name="mode">The ExecutionMode to be used.</param>
        /// <param name="rules">The project invocation rules that are to be invoked.</param>
        public void Run(ExecutionMode mode, ProjectInvocationRule[] rules)
        {
            lock (m_runLock) {
                var d = new ProjectExecutionDialogue();
                PerformActiveAdapterAction(a => a.Run(mode, rules, m_manager, d));
            }
        }

        /// <summary>
        /// Associates the name of a workbook with a WorkbookRuntimeAdapter.
        /// </summary>
        private Dictionary<string, WorkbookRuntimeAdapter> m_workbookAdapters = 
            new Dictionary<string, WorkbookRuntimeAdapter>();

        /// <summary>
        /// Returns the WorkbookRuntimeAdapter associated with a workbook whose name is supplied.  If no 
        /// such adapter exists then one is created.
        /// </summary>
        private WorkbookRuntimeAdapter GetAdapter(string workbookName)
        {
            WorkbookRuntimeAdapter a;
            lock (m_workbookAdapters) {
                if (!m_workbookAdapters.TryGetValue(workbookName, out a)) {
                    a = new WorkbookRuntimeAdapter();
                    m_workbookAdapters[workbookName] = a;
                }
            }
            return a;
        }

        /// <summary>
        /// Removes from the collection of workbook runtime adapters those associated with workbooks that 
        /// are no longer open in the current application.
        /// </summary>
        private void TrimAdapters()
        {
            string[] namesToRemove = GetClosedWorkbookNames();
            foreach (string n in namesToRemove) {
                WorkbookRuntimeAdapter a;
                lock (m_workbookAdapters) {
                    if (m_workbookAdapters.TryGetValue(n, out a)) {
                        m_workbookAdapters.Remove(n);
                    }
                }
                if (a != null) {
                    a.UnloadAll(m_manager);
                }
            }
        }

        /// <summary>
        /// Returns an array of the full names of those workbooks that no longer exist in the current 
        /// application.
        /// </summary>
        private string[] GetClosedWorkbookNames()
        {
            try {
                var workbookNames = m_application.Workbooks.Cast<Excel.Workbook>().Select(w => w.FullName);
                lock (m_workbookAdapters) {
                    return m_workbookAdapters.Select(a => a.Key).Except(workbookNames).ToArray();
                }
            }
            catch (COMException) {
                /* The application may be busy (perhaps because a dialogue is open).  This is not a
                 * problem, we can simply try again later.
                 * */
                return new string[0];
            }
        }

        /// <summary>
        /// Disposes of this ApplicationRuntimeAdapter.
        /// </summary>
        public void Dispose()
        {
            m_timer.Dispose();
            m_manager.Dispose();
        }
    }
}