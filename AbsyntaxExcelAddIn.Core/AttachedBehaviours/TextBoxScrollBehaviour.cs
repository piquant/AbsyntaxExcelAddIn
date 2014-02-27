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

using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// Defines an attached property targeting TextBoxBase instances which, when set to true, causes the
    /// target text box to scroll to the end of its text content when said content changes.
    /// </summary>
    public static class TextBoxScrollBehaviour
    {
        public static readonly DependencyProperty ScrollToEndProperty =
            DependencyProperty.RegisterAttached("ScrollToEnd", typeof(bool), typeof(TextBoxScrollBehaviour),
            new PropertyMetadata(false, OnScrollToEndChanged));

        public static bool GetScrollToEnd(TextBox item)
        {
            return (bool)item.GetValue(ScrollToEndProperty);
        }

        public static void SetScrollToEnd(TextBox item, bool value)
        {
            item.SetValue(ScrollToEndProperty, value);
        }

        private static void OnScrollToEndChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            TextBoxBase item = obj as TextBoxBase;
            if (item == null || !(e.NewValue is bool)) {
                return;
            }
            if ((bool)e.NewValue) {
                item.TextChanged += TextBoxBase_TextChanged;
            }
            else {
                item.TextChanged -= TextBoxBase_TextChanged;
            }
        }

        private static void TextBoxBase_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBoxBase item = sender as TextBoxBase;
            item.ScrollToEnd();
        }
    }
}