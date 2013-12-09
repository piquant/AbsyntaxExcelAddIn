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
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Interaction logic for ExecutionDialogueContent.xaml
    /// </summary>
    public partial class ExecutionDialogueContent : UserControl
    {
        private static readonly DependencyPropertyKey IsExecutingPropertyKey =
            DependencyProperty.RegisterReadOnly("IsExecuting", typeof(bool), typeof(ExecutionDialogueContent), new PropertyMetadata(false));

        public static readonly DependencyProperty IsExecutingProperty = IsExecutingPropertyKey.DependencyProperty;

        private static readonly DependencyPropertyKey ItemsPropertyKey =
            DependencyProperty.RegisterReadOnly("Items", typeof(IExecutionItem[]), typeof(ExecutionDialogueContent), new PropertyMetadata(null));

        public static readonly DependencyProperty ItemsProperty = ItemsPropertyKey.DependencyProperty;

        public ExecutionDialogueContent()
        {
            Helper.EnsureApplicationResources();
            InitializeComponent();
            DataContext = this;
            m_dispatcher = Dispatcher.CurrentDispatcher;
        }

        private Dispatcher m_dispatcher;

        private void Border_GotFocus(object sender, RoutedEventArgs e)
        {
            SetIsSelected(sender, true);
        }

        private void Border_LostFocus(object sender, RoutedEventArgs e)
        {
            SetIsSelected(sender, false);
        }

        private void SetIsSelected(object sender, bool value)
        {
            PerformItemAction(sender as FrameworkElement, r => r.IsSelected = value);
        }

        private static void PerformItemAction(FrameworkElement fe, Action<IExecutionItem> action)
        {
            IExecutionItem item = GetItemFromTag(fe);
            PerformItemAction(item, action);
        }

        private static IExecutionItem GetItemFromTag(FrameworkElement fe)
        {
            return fe == null ? null : fe.Tag as IExecutionItem;
        }

        private static void PerformItemAction(IExecutionItem item, Action<IExecutionItem> action)
        {
            if (item != null) {
                action(item);
            }
        }
        
        public bool IsExecuting
        {
            get { return (bool)GetValue(IsExecutingProperty); }
            private set { Helper.PerformDispatcherAction(m_dispatcher, () => SetValue(IsExecutingPropertyKey, value)); }
        }

        private IRuntimeManager m_manager;

        public IExecutionItem[] Items
        {
            get { return (IExecutionItem[])GetValue(ItemsProperty); }
            private set { Helper.PerformDispatcherAction(m_dispatcher, () => SetValue(ItemsPropertyKey, value)); }
        }

        internal void StartWhenFullyLoaded(ExecutionMode mode, IExecutionItem[] items, IRuntimeManager manager)
        {
            IsExecuting = true;
            m_manager = manager;
            Items = items;
            Action a = () => SelfDisposingBackgroundWorker.RunWorkerAsync((s,e) => Start(mode, items, manager));
            if (IsLoaded) {
                Dispatcher.BeginInvoke(a, DispatcherPriority.ContextIdle, null);
            }
            else {
                RoutedEventHandler handler = null;
                handler = (s, e) => {
                    Loaded -= handler;
                    Dispatcher.BeginInvoke(a, DispatcherPriority.ContextIdle, null);
                };
                Loaded += handler;
            }
        }

        private void Start(ExecutionMode mode, IExecutionItem[] items, IRuntimeManager manager)
        {
            try {
                if (mode == ExecutionMode.Synchronous) {
                    StartSynchronous(items, manager);
                }
                else {
                    StartAsynchronous(items, manager);
                }
            }
            finally {
                IsExecuting = false;
            }
        }

        private void StartSynchronous(IExecutionItem[] items, IRuntimeManager manager)
        {
            using (var mre = new AutoResetEvent(false)) {
                foreach (IExecutionItem item in items) {
                    if (m_abortRequested) break;
                    if (item.State == ProjectExecutionState.Pending) {
                        item.IsSelected = true;
                        item.BeginExecute(manager, i => mre.Set());
                        mre.WaitOne();
                    }
                }
            }
        }

        private void StartAsynchronous(IExecutionItem[] items, IRuntimeManager manager)
        {
            IExecutionItem[] executableItems = items.Where(i => i.State == ProjectExecutionState.Pending).ToArray();
            int count = executableItems.Length;
            if (count > 0) {
                executableItems.First().IsSelected = true;
            }
            using (var mre = new ManualResetEvent(false)) {
                foreach (IExecutionItem item in executableItems) {
                    if (m_abortRequested) break;
                    item.BeginExecute(manager, i => {
                        if (--count == 0) {
                            mre.Set();
                        }
                    });
                }
                mre.WaitOne();
            }
        }

        public event EventHandler Closed;

        protected void OnClosed()
        {
            var handler = Closed;
            if (handler != null) {
                handler(this, EventArgs.Empty);
            }
        }

        private bool m_abortRequested = false;

        private void AbortButton_Click(object sender, RoutedEventArgs e)
        {
            Abort();
        }

        public void Abort()
        {
            Cursor c = Cursor;
            Cursor = Cursors.Wait;
            try {
                m_abortRequested = true;
                IRuntimeManager m = m_manager;
                IExecutionItem[] items = Items;
                if (m == null || items == null) return;
                foreach (IExecutionItem item in items) {
                    item.Abort(m);
                }
            }
            finally {
                Cursor = c;
            }
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            OnClosed();
        }
    }
}