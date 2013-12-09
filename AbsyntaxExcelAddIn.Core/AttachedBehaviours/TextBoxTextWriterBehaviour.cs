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
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// Defines attached properties that facilitate the creation of a TextBoxTextWriter instance to wrap the 
    /// target TextBox and write text to it using TextWriter semantics.
    /// </summary>
    public static class TextBoxTextWriterBehaviour
    {
        public static readonly DependencyProperty AutoCreateProperty =
            DependencyProperty.RegisterAttached("AutoCreate", typeof(bool), typeof(TextBoxTextWriterBehaviour),
            new PropertyMetadata(false, OnAutoCreateChanged));

        public static readonly DependencyProperty TextWriterProperty =
            DependencyProperty.RegisterAttached("TextWriter", typeof(TextWriter), typeof(TextBoxTextWriterBehaviour), 
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnTextWriterChanged));

        /// <summary>
        /// Callback for the AutoCreate attached property's changed event.
        /// </summary>
        /// <remarks>
        /// When the AutoCreate attached property value is changed to true, this signals that a new 
        /// TextBoxTextWriter is to be created to wrap the attached TextBox.  If the attached property is 
        /// bound two-way to a view-model property, the view-model property will be set with the new 
        /// TextBoxTextWriter instance.
        /// </remarks>
        private static void OnAutoCreateChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            TextBox item = obj as TextBox;
            if (item == null || !(e.NewValue is bool)) {
                return;
            }
            if ((bool)e.NewValue) {
                SetTextWriterWhenLoaded(item);
            }
        }

        /// <summary>
        /// Callback for the TextWriter attached property's changed event.
        /// </summary>
        /// <remarks>
        /// This method ensures that an incoming TextBoxTextWriter instance is assigned the attached 
        /// TextBox so that said TextBox will be supplied with text written to the TextBoxTextWriter.
        /// </remarks>
        private static void OnTextWriterChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            TextBox item = obj as TextBox;
            TextBoxTextWriter tw = e.NewValue as TextBoxTextWriter;
            if (item != null && tw != null && tw.Output != item) {
                tw.Output = item;
            }
        }

        /// <summary>
        /// Sets the TextWriter attached property once the associated TextBox has been loaded.
        /// </summary>
        private static void SetTextWriterWhenLoaded(TextBox item)
        {
            Action a = () => {
                TextWriter existing = GetTextWriter(item);
                SetTextWriter(item, new TextBoxTextWriter(item));
                if (existing != null) {
                    existing.Dispose();
                }
            };
            if (item.IsLoaded) {
                a();
            }
            else {
                RoutedEventHandler handler = null;
                handler = (s, e) => {
                    item.Loaded -= handler;
                    a();
                };
                item.Loaded += handler;
            }
        }

        /// <summary>
        /// Returns a value indicating whether a new TextBoxTextWriter is to be created to wrap the 
        /// specified TextBox.
        /// </summary>
        public static bool GetAutoCreate(TextBox item)
        {
            return (bool)item.GetValue(AutoCreateProperty);
        }

        /// <summary>
        /// Sets a value indicating whether a new TextBoxTextWriter is to be created to wrap the 
        /// specified TextBox.
        /// </summary>
        public static void SetAutoCreate(TextBox item, bool value)
        {
            item.SetValue(AutoCreateProperty, value);
        }

        /// <summary>
        /// Gets the TextWriter attached to the specified TextBox.
        /// </summary>
        public static TextWriter GetTextWriter(TextBox item)
        {
            return (TextWriter)item.GetValue(TextWriterProperty);
        }

        /// <summary>
        /// Sets the TextWriter attached to the specified TextBox.
        /// </summary>
        public static void SetTextWriter(TextBox item, TextWriter value)
        {
            item.SetValue(TextWriterProperty, value);
        }
    }
}