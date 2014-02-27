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
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using lex = System.Linq.Expressions;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// Defines attached properties targeting ComboBox objects which allow properties of the 
    /// underlying TextBox to be set (i.e. the TextBox that is used when IsEditable = true).
    /// </summary>
    public static class EditableComboBoxBehaviour
    {
        public static readonly DependencyProperty MaxLengthProperty =
            DependencyProperty.RegisterAttached("MaxLength", typeof(int), typeof(EditableComboBoxBehaviour),
            new PropertyMetadata(OnMaxLengthChanged));

        public static readonly DependencyProperty CharacterCasingProperty =
            DependencyProperty.RegisterAttached("CharacterCasing", typeof(CharacterCasing), typeof(EditableComboBoxBehaviour),
            new PropertyMetadata(OnCharacterCasingChanged));

        #region MaxLength

        public static int GetMaxLength(ComboBox item)
        {
            return (int)item.GetValue(MaxLengthProperty);
        }

        public static void SetMaxLength(ComboBox item, int value)
        {
            item.SetValue(MaxLengthProperty, value);
        }

        private static void OnMaxLengthChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            OnPropertyChanged<int>(obj as ComboBox, e, tb => tb.MaxLength);
        }

        #endregion

        #region CharacterCasing

        public static CharacterCasing GetCharacterCasing(ComboBox item)
        {
            return (CharacterCasing)item.GetValue(CharacterCasingProperty);
        }

        public static void SetCharacterCasing(ComboBox item, CharacterCasing value)
        {
            item.SetValue(CharacterCasingProperty, value);
        }

        private static void OnCharacterCasingChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            OnPropertyChanged<CharacterCasing>(obj as ComboBox, e, tb => tb.CharacterCasing);
        }

        #endregion

        private static void OnPropertyChanged<T>(ComboBox item, DependencyPropertyChangedEventArgs e, lex.Expression<Func<TextBox, T>> property)
        {
            if (item == null || !(e.NewValue is T)) return;
            if (item.IsLoaded) {
                FindTextBoxAndSetProperty(item, property, (T)e.NewValue);
            }
            else {
                RoutedEventHandler handler = null;
                handler = new RoutedEventHandler((o, a) => {
                    FindTextBoxAndSetProperty(item, property, (T)e.NewValue);
                    item.Loaded -= handler;
                });
                item.Loaded += handler;
            }
        }

        private static void FindTextBoxAndSetProperty<T>(ComboBox item, lex.Expression<Func<TextBox, T>> property, T value)
        {
            var textBox = item.FindChild<TextBox>("PART_EditableTextBox");
            if (textBox != null) {
                SetProperty(textBox, property, value);
            }
        }

        private static void SetProperty<T>(TextBox tb, lex.Expression<Func<TextBox, T>> property, T value)
        {
            var member = property.Body as lex.MemberExpression;
            var pi = member.Member as PropertyInfo;
            MethodInfo mi = pi.GetSetMethod();
            lex.ParameterExpression tbpe = lex.Expression.Parameter(typeof(TextBox));
            lex.ParameterExpression vpe = lex.Expression.Parameter(typeof(T));
            lex.MethodCallExpression mce = lex.Expression.Call(tbpe, mi, vpe);
            lex.Expression<Action<TextBox, T>> action = lex.Expression.Lambda<Action<TextBox, T>>(mce, tbpe, vpe);
            action.Compile()(tb, value);
        }
    }
}