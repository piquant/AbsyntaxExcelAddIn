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
using System.IO;
using System.Linq;
using System.Reflection;
using AbsyntaxExcelAddIn.Resources;
using MI2.FrameworkAdapter;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// An implementation of the IExecutionItem interface that acts both as a UI view-model and a 
    /// facilitator for invoking Absyntax projects.
    /// </summary>
    internal sealed class ExecutionItem : NotifyPropertyChangedBase, IExecutionItem
    {
        /// <summary>
        /// Initialises a new ExecutionItem instance.
        /// </summary>
        /// <param name="rule">An IProjectInvocationRule instance containing up-to-date details of the 
        /// project to be executed.</param>
        /// <param name="detail">A ProjectExecutionDetail instance containing details of the last time the 
        /// associated project invocation rule was executed.</param>
        public ExecutionItem(IProjectInvocationRule rule, ProjectExecutionDetail detail)
        {
            m_rule = rule;
            Detail = detail;
            Log = detail.Log;
            AutoCreateLog = detail.Log == null;
            State = rule.CanExecute ? ProjectExecutionState.Pending : ProjectExecutionState.Ineligible;
        }

        private IProjectInvocationRule m_rule;

        /// <summary>
        /// Gets a ProjectExecutionDetail instance containing details of the last time the associated 
        /// project invocation rule was executed.
        /// </summary>
        /// <remarks>
        /// This instance is typically updated during project invocation.
        /// </remarks>
        public ProjectExecutionDetail Detail { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this ExecutionRule can be executed.
        /// </summary>
        public bool CanExecute
        {
            get { return m_rule.CanExecute; }
        }

        private bool m_isSelected;

        /// <summary>
        /// Gets or sets a value indicating whether this ExecutionItem is selected in a list.
        /// </summary>
        public bool IsSelected
        {
            get { return m_isSelected; }
            set { SetProperty(ref m_isSelected, value, () => IsSelected); }
        }

        /// <summary>
        /// Gets the identifier assigned to the project invocation rule represented by this ExecutionItem.
        /// </summary>
        public int Id
        {
            get { return m_rule.Id; }
        }

        /// <summary>
        /// Gets the identifier assigned by the Absyntax runtime to a project when it was loaded.  A null 
        /// value indicates that the project has not yet been loaded, has not been loaded successfully or
        /// has been unloaded after an invocation.
        /// </summary>
        public int? Key
        {
            get { return Detail.Key; }
            private set { Detail.Key = value; }
        }

        /// <summary>
        /// Gets the path of the project to be executed or, where no such path is defined by the underlying 
        /// rule, a placeholder description.
        /// </summary>
        public string ProjectPath
        {
            get { return String.IsNullOrWhiteSpace(m_rule.ProjectPath) ? TextResources.ProjectPathNotSpecified : m_rule.ProjectPath; }
        }

        private ProjectExecutionState m_state;

        /// <summary>
        /// Gets the execution state of this ExecutionItem.
        /// </summary>
        public ProjectExecutionState State
        {
            get { return m_state; }
            private set {
                if (SetProperty(ref m_state, value, () => State)) {
                    OnPropertyChanged(() => IsExecuting);
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether a project invocation is currently in progress.
        /// </summary>
        public bool IsExecuting
        {
            get { return State == ProjectExecutionState.Executing; }
        }

        private bool m_autoCreateLog;

        /// <summary>
        /// Gets a value indicating whether a TextWriter used for log messages is to be created 
        /// automatically.
        /// </summary>
        public bool AutoCreateLog
        {
            get { return m_autoCreateLog; }
            private set { SetProperty(ref m_autoCreateLog, value, () => AutoCreateLog); }
        }

        private TextWriter m_log;

        /// <summary>
        /// Gets or sets the TextWriter to be used by the Absyntax framework to write project runtime 
        /// messages.
        /// </summary>
        public TextWriter Log
        {
            get { return m_log; }
            set {
                SetProperty(ref m_log, value, () => Log);
                Detail.Log = value;
            }
        }

        private Func<IOperationResult> m_func;

        /// <summary>
        /// Starts an asynchronous execution of the represented Absyntax project.
        /// </summary>
        /// <param name="manager">The IRuntimeManager required to perform the underlying project runtime 
        /// tasks.</param>
        /// <param name="callback">The action to be invoked upon completion of the task.</param>
        /// <exception cref="System.InvalidOperationException">The ExecutionItem is not in a valid state
        /// to begin execution.</exception>
        public void BeginExecute(IRuntimeManager manager, Action<IExecutionItem> callback)
        {
            CheckState();
            State = ProjectExecutionState.Executing;
            m_func = () => LoadAndInvoke(manager);
            m_func.BeginInvoke(EndExecute, callback);
        }

        /// <summary>
        /// Checks whether this ExecutionItem is pending and throws a <see cref="System.InvalidOperationException"/>
        /// if it is not.
        /// </summary>
        private void CheckState()
        {
            if (State != ProjectExecutionState.Pending) {
                throw new InvalidOperationException();
            }
        }

        /// <summary>
        /// Attempts to load the target project if it is not already loaded, and then invoke it.
        /// </summary>
        private IOperationResult LoadAndInvoke(IRuntimeManager manager)
        {
            LoadProjectIfNecessary(manager);
            if (m_aborted) return null;
            IOperationResult result = null;
            try {
                if (m_rule.UsesInput) {
                    IEnumerable<object> data = m_rule.ReadInputData();
                    result = Invoke(manager, data.ToArray());
                }
                else {
                    result = manager.Invoke(Key.Value);
                }
            }
            catch (InvalidOperationException) {
                if (!m_aborted) throw;
            }
            return m_aborted ? null : result;
        }

        /// <summary>
        /// Ends an asynchronous project invocation.
        /// </summary>
        private void EndExecute(IAsyncResult result)
        {
            int? key = Key;
            Key = null;
            try {
                IOperationResult or = m_func.EndInvoke(result);
                if (or == null) {
                    State = ProjectExecutionState.Aborted;
                }
                else {
                    OperationStatus s = or.Status;
                    switch (s) {
                        case OperationStatus.Ok:
                            Key = key;
                            State = ProjectExecutionState.Completed;
                            if (m_rule.UsesOutput) {
                                WriteOutputData(or.Data);
                            }
                            break;
                        case OperationStatus.Aborted:
                            State = ProjectExecutionState.Aborted;
                            break;
                        case OperationStatus.ServiceTimedOut:
                            State = ProjectExecutionState.TimedOut;
                            break;
                        default:
                            State = ProjectExecutionState.Errors;
                            break;
                    }
                }
            }
            catch (TimeoutException ex) {
                State = ProjectExecutionState.TimedOut;
                WriteLineToLog(ex.Message);
            }
            catch (Exception ex) {
                State = ProjectExecutionState.Errors;
                WriteLineToLog(ex.Message);
            }
            Action<IExecutionItem> callback = (Action<IExecutionItem>)result.AsyncState;
            if (callback != null) {
                callback(this);
            }
        }

        /// <summary>
        /// Attempts to write an object to a range of cells defined by the various output-related state 
        /// values of the encapsulated ProjectInvocationRule.  In the event of an error being thrown, the 
        /// error message is written to the output and the execution state of this ExecutionItem is set 
        /// to ProjectExecutionState.WriteDataErrors.
        /// </summary>
        private void WriteOutputData(object data)
        {
            try {
                m_rule.WriteOutputData(data);
            }
            catch {
                string orn = String.Format("{0}!{1}", m_rule.OutputSheetName, m_rule.OutputCellRange);
                string msg = String.Format(TextResources.Msg_FailedToWriteOutputData, orn);
                WriteLineToLog(msg);
                State = ProjectExecutionState.WriteDataErrors;
            }
        }

        /// <summary>
        /// Invokes a project to which input data is to be passed.
        /// </summary>
        private IOperationResult Invoke(IRuntimeManager manager, object[] items)
        {
            IOperationResult result;
            switch (Detail.InputDataRequirement) {
                case DataRequirement.None:
                case DataRequirement.SingleValue:
                    result = Invoke(manager, items, DataRequirement.SingleValue, DataRequirement.StronglyTypedArray, DataRequirement.ObjectArray);
                    break;
                case DataRequirement.StronglyTypedArray:
                    result = Invoke(manager, items, DataRequirement.StronglyTypedArray, DataRequirement.SingleValue, DataRequirement.ObjectArray);
                    break;
                default: // ObjectArray
                    result = Invoke(manager, items, DataRequirement.ObjectArray, DataRequirement.SingleValue, DataRequirement.StronglyTypedArray);
                    break;
            }
            return result;
        }

        /// <summary>
        /// Invokes a project to which input data is to be passed.  If there is a mismatch between the 
        /// supplied data and the input data requirements of the project, this method may make multiple
        /// attempts to invoke the project, using different forms of data on each attempt.
        /// </summary>
        /// <remarks>
        /// Absyntax does not report the input data requirements of a project loaded into the runtime.
        /// This is because users are reasonably expected to know these details.  Nonetheless, the runtime 
        /// API provides for a single, serialisable <see cref="System.Object"/> to be supplied as input 
        /// data for an invocation and this causes problems if the object's underlying type cannot be 
        /// converted to that required by the project.  Specifically, Absyntax will report a startup data 
        /// conversion problem, unload the project and close the communication channel between the client 
        /// and the hosting service.
        /// <para />
        /// Project-loading is a relatively slow process and it is desirable to do this as few times as 
        /// possible.  For this reason, the specific input-data aggregation method of a successful 
        /// invocation is cached so that it can be used first on the next invocation, the presumption 
        /// being that this method will be the one most likely to succeed next time around.  Of course, 
        /// there are no guarantees because the input data may change in a way that invalidates the previous 
        /// method, or the invocation rule's target project may have changed completely.
        /// <para />
        /// If an attempt to invoke a loaded project with startup data fails, fallback data aggregation 
        /// methods are used.  Because such failure results in the project being unloaded (as described 
        /// above), the project is reloaded prior to each subsequent invocation attempt.
        /// </remarks>
        /// <param name="manager">The IRuntimeManager required to perform the underlying project-load and
        /// invocation operations.</param>
        /// <param name="items">An array containing the values of Excel worksheet cells in the associated 
        /// project invocation rule's defined input range.</param>
        /// <param name="requirements">A sequence of ways in which data is to be presented to the project
        /// upon invocation in the event of a conversion problem during an invocation.</param>
        /// <returns>An <see cref="MI2.FrameworkAdapter.IOperationResult"/> encapsulating details of the 
        /// invocation.</returns>
        private IOperationResult Invoke(IRuntimeManager manager, object[] items, params DataRequirement[] requirements)
        {
            IOperationResult result = null;
            foreach (DataRequirement dr in requirements) {
                object data;
                switch (dr) {
                    case DataRequirement.SingleValue:
                        if (items.Length > 1) continue;
                        data = items.First();
                        break;
                    case DataRequirement.StronglyTypedArray:
                        if (!TryGetStronglyTypedArray(items, out data)) continue;
                        break;
                    default:
                        data = items;
                        break;
                }
                if (m_aborted) break;
                result = manager.Invoke(Key.Value, data);
                switch (result.Status) {
                    case OperationStatus.Ok:
                        Detail.InputDataRequirement = dr;
                        return result;
                    case OperationStatus.StartupDataConversionProblem:
                        Key = null;
                        if (dr != requirements.Last()) {
                            LoadProjectIfNecessary(manager);
                        }
                        break;
                    default:
                        Key = null;
                        return result;
                }
                if (m_aborted) break;
            }
            Detail.InputDataRequirement = DataRequirement.None;
            return result;
        }

        /// <summary>
        /// Flags whether an abort request has been received.
        /// </summary>
        private volatile bool m_aborted;

        private object m_abortLock = new object();

        /// <summary>
        /// Invokes the supplied action if an abort request has not been received.
        /// </summary>
        private void PerformNonAbortedAction(Action action)
        {
            if (!m_aborted) {
                lock (m_abortLock) {
                    if (!m_aborted) {
                        action();
                    }
                }
            }
        }

        /// <summary>
        /// Attempts to convert the supplied object array into a strongly-typed array.
        /// </summary>
        /// <param name="items">The array of items to convert.</param>
        /// <param name="data">The converted, strongly-typed array or null if no such conversion is 
        /// possible.</param>
        /// <returns>True if there is at least one non-null item and that all non-null items share the
        /// same type.</returns>
        private bool TryGetStronglyTypedArray(object[] items, out object data)
        {
            data = null;
            bool hasRefTypes = items.Any(i => i != null && !i.GetType().IsValueType);
            bool hasValTypes = items.Any(i => i != null && i.GetType().IsValueType);
            if ((hasRefTypes && hasValTypes) || (!hasRefTypes && !hasValTypes)) {
                return false;
            }
            Type firstType = items.First(i => i != null).GetType();
            bool multipleTypes = items.Any(i => i != null && i.GetType() != firstType);
            if (multipleTypes) {
                return false;
            }
            data = CreateArray(items, firstType);
            return true;
        }

        private static readonly MethodInfo s_castMethod =
            typeof(ExecutionItem).GetMethod("CreateArray", BindingFlags.NonPublic | BindingFlags.Static, null, new Type[] { typeof(object[]) }, null);

        /// <summary>
        /// Casts the supplied object array to an array of items of the specified type.
        /// </summary>
        private static object CreateArray(object[] items, Type toType)
        {
            MethodInfo mi = s_castMethod.MakeGenericMethod(toType);
            return mi.Invoke(null, new object[] { items });
        }

        /// <summary>
        /// Casts the supplied object array to an array of items of type T.
        /// </summary>
        private static T[] CreateArray<T>(object[] items)
        {
            if (typeof(T).IsValueType) {
                return items.Where(i => i != null).Cast<T>().ToArray();
            }
            return items.Cast<T>().ToArray();
        }

        /// <summary>
        /// Loads the Absyntax project identified by the underlying project invocation rule, but only 
        /// if it has not already been loaded.
        /// </summary>
        private void LoadProjectIfNecessary(IRuntimeManager manager)
        {
            PerformNonAbortedAction(() => {
                if (Key == null) {
                    IStartupArgs args = CreateArgs();
                    WriteLineToLog(TextResources.Msg_LoadingProject);
                    Key = manager.Load(m_rule.ProjectPath, args);
                }
            });
        }

        /// <summary>
        /// Creates startup arguments in readiness for loading an Absyntax project.
        /// </summary>
        private IStartupArgs CreateArgs()
        {
            return new StartupArgs() {
                Log = this.Log, 
                OperationTimeout = Helper.GetMilliseconds(m_rule.TimeLimit, m_rule.Unit)
            };
        }

        /// <summary>
        /// Aborts a current execution.
        /// </summary>
        /// <param name="manager">The IRuntimeManager required to perform the underlying abort operation.</param>
        public void Abort(IRuntimeManager manager)
        {
            PerformNonAbortedAction(() => {
                m_aborted = true;
                if (IsExecuting) {
                    manager.Unload(Key.Value);
                }
            });
        }

        /// <summary>
        /// Writes the supplied string to the log.
        /// </summary>
        private void WriteLineToLog(string text)
        {
            PerformLogAction(l => l.WriteLine(text));
        }

        /// <summary>
        /// Invokes the supplied action if the logging TextWriter is not null.
        /// </summary>
        private void PerformLogAction(Action<TextWriter> action)
        {
            TextWriter log = Log;
            if (log != null) {
                action(log);
            }
        }
    }
}