using System;
using System.Windows;
using System.Windows.Controls;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// Defines an attached property targeting ListBoxItems which, when set to true, causes the target 
    /// item to be scrolled into view if it is selected.
    /// </summary>
    /// <remarks>
    /// Courtesy of Josh Smith: http://www.codeproject.com/KB/WPF/AttachedBehaviors.aspx and licensed
    /// under the Code Project Open License (CPOL) (http://www.codeproject.com/info/cpol10.aspx).
    /// </remarks>
    public static class ListBoxItemBehaviour
    {
        public static readonly DependencyProperty IsBroughtIntoViewWhenSelectedProperty =
            DependencyProperty.RegisterAttached("IsBroughtIntoViewWhenSelected", typeof(bool), typeof(ListBoxItemBehaviour),
            new PropertyMetadata(false, OnIsBroughtIntoViewWhenSelectedChanged));

        public static bool GetIsBroughtIntoViewWhenSelected(ListBoxItem item)
        {
            return (bool)item.GetValue(IsBroughtIntoViewWhenSelectedProperty);
        }
        
        public static void SetIsBroughtIntoViewWhenSelected(ListBoxItem item, bool value)
        {
            item.SetValue(IsBroughtIntoViewWhenSelectedProperty, value);
        }
        
        private static void OnIsBroughtIntoViewWhenSelectedChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            ListBoxItem item = obj as ListBoxItem;
            if (item == null || !(e.NewValue is bool)) {
                return;
            }
            if ((bool)e.NewValue) {
                item.Selected += OnItemSelected;
            }
            else {
                item.Selected -= OnItemSelected;
            }
            if (item.IsSelected) {
                item.BringIntoView();
            }
        }

        private static void OnItemSelected(object sender, RoutedEventArgs e)
        {
            /* Only react to the Selected event raised by the ListBoxItem whose IsSelected property was 
             * modified.  Ignore all ancestors who are merely reporting that a descendant's Selected fired.
             * */
            if (Object.ReferenceEquals(sender, e.OriginalSource)) {
                var item = e.OriginalSource as ListBoxItem;
                if (item != null) {
                    item.BringIntoView();
                }
            }
        }
    }
}