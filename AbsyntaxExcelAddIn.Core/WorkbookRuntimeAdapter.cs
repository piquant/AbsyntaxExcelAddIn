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

using System.Collections.Generic;
using System.Linq;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Manages a set of isolated project runtime "slots" for a collection of project invocation rules.
    /// </summary>
    /// <remarks>
    /// Absyntax projects are executed in accordance with project invocation rules that are defined by the 
    /// add-in user.  Collections of rules may be defined on a per-workbook basis.  It is the responsibility 
    /// of this class to coordinate an execution cycle for such a collection of rules.
    /// </remarks>
    public sealed class WorkbookRuntimeAdapter
    {
        /// <summary>
        /// Initialises a new WorkbookRuntimeAdapter instance.
        /// </summary>
        public WorkbookRuntimeAdapter()
        { }

        /// <summary>
        /// Associates unique Absyntax IsolatedRuntimeAdapter project "slot" numbers with a sets of details 
        /// representing a project's last execution.
        /// </summary>
        private Dictionary<int, ProjectExecutionDetail> m_ruleState = new Dictionary<int, ProjectExecutionDetail>();

        /// <summary>
        /// Re-bases the unique identifiers assigned to a set of ProjectInvocationRule instances.
        /// </summary>
        /// <param name="rules">An array of ProjectInvocationRules whose identifiers are to be rationalised.</param>
        public void RationaliseIds(ProjectInvocationRule[] rules)
        {
            IList<int> list = GetImmutableIds(rules);
            var rulesToUpdate = rules.Where(r => !list.Contains(r.Id)).ToArray();
            foreach (ProjectInvocationRule rule in rulesToUpdate) {
                int id = Helper.CreateId(list);
                rule.Id = id;
                list.Add(id);
            }
        }

        /// <summary>
        /// Returns a list of ProjectInvocationRule ids that cannot be changed.
        /// </summary>
        /// <remarks>
        /// A ProjectInvocationRule's id cannot be changed if this WorkbookRuntimeAdapter is already
        /// managing an execution slot for it.
        /// </remarks>
        private IList<int> GetImmutableIds(ProjectInvocationRule[] rules)
        {
            lock (m_ruleState) {
                return new List<int>(rules.Select(r => r.Id).Intersect(m_ruleState.Keys));
            }
        }

        /// <summary>
        /// Invokes those of the supplied project invocation rules that are valid and enabled.
        /// </summary>
        /// <param name="mode">The ExecutionMode to be used.</param>
        /// <param name="rules">The project invocation rules that are to be invoked.</param>
        /// <param name="manager">The IRuntimeManager with which to perform the underlying Absyntax project-
        /// loading and invocation activities.</param>
        /// <param name="coordinator">The IExecutionCoordinator needed to coordinate the execution sequence.</param>
        public void Run(ExecutionMode mode, ProjectInvocationRule[] rules, IRuntimeManager manager, IExecutionCoordinator coordinator)
        {
            UnloadProjectsWhereNecessary(rules, manager);
            ExecutionItem[] items = CreateExecutionItems(rules);
            coordinator.Start(mode, items, manager);
            lock (m_ruleState) {
                foreach (ExecutionItem item in items) {
                    switch (item.State) {
                        case ProjectExecutionState.Completed:
                        case ProjectExecutionState.WriteDataErrors:
                            if (!m_ruleState.ContainsKey(item.Id)) {
                                m_ruleState[item.Id] = item.Detail;
                            }
                            break;
                        case ProjectExecutionState.Ineligible:
                        case ProjectExecutionState.Pending:
                            // Do nothing
                            break;
                        default:
                            m_ruleState.Remove(item.Id);
                            break;
                    }
                    ProjectInvocationRule rule = rules.First(r => r.Id == item.Id);
                    rule.LastExecutionResult = GetLastExecutionResult(item.State);
                }
            }
        }

        /// <summary>
        /// Converts a ProjectExecutionState into an ExecutionResult.
        /// </summary>
        private static ExecutionResult GetLastExecutionResult(ProjectExecutionState state)
        {
            ExecutionResult result;
            switch (state) {
                case ProjectExecutionState.Completed:
                    result = ExecutionResult.Ok;
                    break;
                case ProjectExecutionState.Ineligible:
                case ProjectExecutionState.Pending:
                    result = ExecutionResult.Unknown;
                    break;
                default:
                    result = ExecutionResult.Errors;
                    break;
            }
            return result;
        }

        /// <summary>
        /// Instructs the Absyntax runtime host to close certain project slots.
        /// </summary>
        /// <remarks>
        /// Project slots to be closed are:
        /// (a) those that are not referred to in the incoming invocation rules;
        /// (b) those that are associated with rules whose material details have changed;
        /// (c) those that are associated with rules for which the property ReloadProjectBeforeExecuting 
        /// returns true.
        /// </remarks>
        private void UnloadProjectsWhereNecessary(IProjectInvocationRule[] rules, IRuntimeManager manager)
        {
            KeyValuePair<int, ProjectExecutionDetail>[] items;
            lock (m_ruleState) {
                items = m_ruleState.ToArray();
            }
            var list = new List<int>();
            foreach (var kvp in items) {
                int ruleId = kvp.Key;
                ProjectExecutionDetail ped = kvp.Value;
                IProjectInvocationRule matchingRule = rules.FirstOrDefault(r => r.Id == ruleId);
                if (matchingRule == null || DetailsHaveChanged(ped, matchingRule) || matchingRule.ReloadProjectBeforeExecuting) {
                    int? key = ped.Key;
                    if (key.HasValue) {
                        manager.Unload(key.Value);
                    }
                    list.Add(ruleId);
                }
            }
            lock (m_ruleState) {
                list.ForEach(i => m_ruleState.Remove(i));
            }
        }

        /// <summary>
        /// Identifies whether material information differs between a ProjectExecutionDetail and an
        /// IProjectInvocationRule.
        /// </summary>
        private bool DetailsHaveChanged(ProjectExecutionDetail detail, IProjectInvocationRule rule)
        {
            int ms1 = Helper.GetMilliseconds(detail.TimeLimit, detail.Unit);
            int ms2 = Helper.GetMilliseconds(rule.TimeLimit, rule.Unit);
            return ms1 != ms2 || detail.Path != rule.ProjectPath;
        }

        /// <summary>
        /// Creates a new ExecutionItem instance for each supplied IProjectInvocationRule.
        /// </summary>
        private ExecutionItem[] CreateExecutionItems(IProjectInvocationRule[] rules)
        {
            var list = new List<ExecutionItem>();
            lock (m_ruleState) {
                foreach (IProjectInvocationRule rule in rules) {
                    ProjectExecutionDetail detail;
                    if (!m_ruleState.TryGetValue(rule.Id, out detail)) {
                        detail = ProjectExecutionDetail.Create(rule);
                    }
                    var item = new ExecutionItem(rule, detail);
                    list.Add(item);
                }
            }
            return list.ToArray();
        }

        /// <summary>
        /// Causes this WorkbookRuntimeAdapter to close all open project runtime slots.
        /// </summary>
        /// <param name="manager">The IRuntimeManager to be used to perform the unload.</param>
        public void UnloadAll(IRuntimeManager manager)
        {
            ProjectExecutionDetail[] items;
            lock (m_ruleState) {
                items = m_ruleState.Values.ToArray();
                m_ruleState.Clear();
            }
            foreach (var key in items.Select(i => i.Key)) {
                if (key.HasValue) {
                    manager.Unload(key.Value);
                }
            }
        }
    }
}