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
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using AbsyntaxExcelAddIn.Core.Converters;
using AbsyntaxExcelAddIn.Resources;
using Microsoft.Win32;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Interaction logic for ConfigurationDialogueContent.xaml
    /// </summary>
    public partial class ConfigurationDialogueContent : UserControl
    {
        public static readonly DependencyProperty ModeProperty =
            DependencyProperty.Register("Mode", typeof(ExecutionMode), typeof(ConfigurationDialogueContent));

        public static readonly DependencyProperty SelectedRuleProperty =
            DependencyProperty.Register("SelectedRule", typeof(ProjectInvocationRule), typeof(ConfigurationDialogueContent), new PropertyMetadata(OnSelectedRuleChanged));

        private static readonly DependencyPropertyKey CanDemoteSelectedRulePropertyKey =
            DependencyProperty.RegisterReadOnly("CanDemoteSelectedRule", typeof(bool), typeof(ConfigurationDialogueContent), new PropertyMetadata());

        public static readonly DependencyProperty CanDemoteSelectedRuleProperty = CanDemoteSelectedRulePropertyKey.DependencyProperty;

        private static readonly DependencyPropertyKey CanPromoteSelectedRulePropertyKey =
            DependencyProperty.RegisterReadOnly("CanPromoteSelectedRule", typeof(bool), typeof(ConfigurationDialogueContent), new PropertyMetadata());

        public static readonly DependencyProperty CanPromoteSelectedRuleProperty = CanPromoteSelectedRulePropertyKey.DependencyProperty;

        private static void OnSelectedRuleChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            ConfigurationDialogueContent c = (ConfigurationDialogueContent)obj;
            c.UpdateCanDemoteSelectedRule();
            c.UpdateCanPromoteSelectedRule();
        }

        public ConfigurationDialogueContent()
        {
            Helper.EnsureApplicationResources();
            InitializeComponent();
            DataContext = this;
            var modes = Enum.GetValues(typeof(ExecutionMode)).Cast<ExecutionMode>();
            var c = new ExecutionModeConverter();
            m_availableExecutionModeNames = modes.Select(o => (string)c.Convert(o, null, null, null));
            Mode = ExecutionMode.Synchronous;
        }

        public IWorksheetProvider WorksheetProvider { get; set; }

        private void InvokeWorksheetProviderAction(Action<IWorksheetProvider> action)
        {
            var wp = WorksheetProvider;
            if (wp != null) {
                action(wp);
            }
        }

        private readonly IEnumerable<string> m_availableExecutionModeNames;

        public IEnumerable<string> AvailableExecutionModeNames
        {
            get { return m_availableExecutionModeNames; }
        }

        private ObservableCollection<ProjectInvocationRule> m_rules = new ObservableCollection<ProjectInvocationRule>();

        public ObservableCollection<ProjectInvocationRule> Rules
        {
            get { return m_rules; }
        }

        public void SetRules(ProjectInvocationRule[] rules)
        {
            m_rules.Clear();
            if (rules != null) {
                foreach (ProjectInvocationRule rule in rules) {
                    Add(rule);
                }
                SelectedRule = rules.FirstOrDefault();
            }
        }

        public ProjectInvocationRule[] GetRules()
        {
            return m_rules.ToArray();
        }

        public event EventHandler Accepted;

        public event EventHandler Cancelled;

        private void OnEvent(EventHandler handler)
        {
            if (handler != null) {
                handler(this, EventArgs.Empty);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            OnEvent(Accepted);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            OnEvent(Cancelled);
        }

        public ExecutionMode Mode
        {
            get { return (ExecutionMode)GetValue(ModeProperty); }
            set { SetValue(ModeProperty, value); }
        }

        public ProjectInvocationRule SelectedRule
        {
            get { return (ProjectInvocationRule)GetValue(SelectedRuleProperty); }
            set { SetValue(SelectedRuleProperty, value); }
        }

        private void UpdateCanDemoteSelectedRule()
        {
            ProjectInvocationRule r = SelectedRule;
            CanDemoteSelectedRule = r != null && r != m_rules.LastOrDefault();
        }

        public bool CanDemoteSelectedRule
        {
            get { return (bool)GetValue(CanDemoteSelectedRuleProperty); }
            private set { SetValue(CanDemoteSelectedRulePropertyKey, value); }
        }

        private void UpdateCanPromoteSelectedRule()
        {
            ProjectInvocationRule r = SelectedRule;
            CanPromoteSelectedRule = r != null && r != m_rules.FirstOrDefault();
        }

        public bool CanPromoteSelectedRule
        {
            get { return (bool)GetValue(CanPromoteSelectedRuleProperty); }
            private set { SetValue(CanPromoteSelectedRulePropertyKey, value); }
        }

        private void Border_GotFocus(object sender, RoutedEventArgs e)
        {
            SetIsSelected(sender, true);
        }

        private void Border_LostFocus(object sender, RoutedEventArgs e)
        {
            SetIsSelected(sender, false);
        }

        private void Border_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Border b = (Border)sender;
            UIElement uie = (UIElement)LogicalTreeHelper.FindLogicalNode(b, "PP");
            Keyboard.Focus(uie);
        }

        private void SetIsSelected(object sender, bool value)
        {
            PerformRuleAction(sender as FrameworkElement, r => r.IsSelected = value);
        }

        private static void PerformRuleAction(FrameworkElement fe, Action<ProjectInvocationRule> action)
        {
            ProjectInvocationRule pir = GetRuleFromTag(fe);
            PerformRuleAction(pir, action);
        }

        private static void PerformRuleAction(ProjectInvocationRule pir, Action<ProjectInvocationRule> action)
        {
            if (pir != null) {
                action(pir);
            }
        }

        private static ProjectInvocationRule GetRuleFromTag(FrameworkElement fe)
        {
            return fe == null ? null : fe.Tag as ProjectInvocationRule;
        }

        private static readonly string s_projectExt = "apj";

        private void BrowseProjectsButton_Click(object sender, RoutedEventArgs e)
        {
            ProjectInvocationRule pir = GetRuleFromTag(sender as FrameworkElement);
            string path = null;
            PerformRuleAction(pir, r => path = r.ProjectPath);
            var d = new OpenFileDialog();
            d.DefaultExt = s_projectExt;
            try {
                d.InitialDirectory = System.IO.Path.GetDirectoryName(path);
                d.FileName = System.IO.Path.GetFileName(path);
            }
            catch { }
            d.Filter = String.Format("{0}|*.{1}", TextResources.AbsyntaxProjects, s_projectExt);
            if (d.ShowDialog() == true) {
                PerformRuleAction(pir, r => r.ProjectPath = d.FileName);
            }
        }

        private void AddRuleButton_Click(object sender, RoutedEventArgs e)
        {
            InvokeWorksheetProviderAction(wp => {
                int id = Helper.CreateId(m_rules.Select(r => r.Id));
                var pir = new ProjectInvocationRule(wp, id);
                Add(pir);
                pir.IsSelected = true;
            });
        }

        private void Add(ProjectInvocationRule rule)
        {
            m_rules.Add(rule);
        }

        private void RemoveRuleButton_Click(object sender, RoutedEventArgs e)
        {
            ProjectInvocationRule r = SelectedRule;
            int index = m_rules.IndexOf(r);
            m_rules.RemoveAt(index);
            int count = m_rules.Count;
            if (index < count) {
                m_rules[index].IsSelected = true;
            }
            else if (count > 0) {
                m_rules.Last().IsSelected = true;
            }
        }

        private void DemoteRuleButton_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedRule(i => i + 1);
        }

        private void PromoteRuleButton_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedRule(i => i - 1);
        }

        private void MoveSelectedRule(Func<int, int> indexModifier)
        {
            ProjectInvocationRule r = SelectedRule;
            int index = m_rules.IndexOf(r);
            m_rules.RemoveAt(index);
            m_rules.Insert(indexModifier(index), r);
        }
    }
}