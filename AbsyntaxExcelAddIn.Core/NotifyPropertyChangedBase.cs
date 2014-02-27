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
using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Abstract base implementation of the <see cref="System.ComponentModel.INotifyPropertyChanged"/> 
    /// interface.
    /// </summary>
    public abstract class NotifyPropertyChangedBase : INotifyPropertyChanged
    {
        protected NotifyPropertyChangedBase()
        { }

        /// <summary>
        /// The event that is raised when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises the PropertyChanged event.
        /// </summary>
        /// <param name="property">A lambda expression that evaluates to the PropertyInfo of a property 
        /// whose value has changed.</param>
        protected void OnPropertyChanged<T>(Expression<Func<T>> property)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) {
                var propertyName = GetPropertyName(property);
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        /// <summary>
        /// Returns the name of the .NET property whose lambda expression is supplied.
        /// </summary>
        /// <param name="property">A lambda expression that evaluates to the PropertyInfo of a property.</param>
        /// <returns>The property name.</returns>
        private static string GetPropertyName<T>(Expression<Func<T>> property)
        {
            var propertyInfo = ((MemberExpression)property.Body).Member as PropertyInfo;
            if (propertyInfo == null) {
                throw new ArgumentException("property is not a .NET property");
            }
            return propertyInfo.Name;
        }

        /// <summary>
        /// If newValue is not equal to underlyingValue, updates underlyingValue to be equal to newValue 
        /// and raises the PropertyChanged event for the member represented by property.
        /// </summary>
        /// <typeparam name="T">The Type of data being updated</typeparam>
        /// <param name="underlyingValue">A reference to the backing variable for property.</param>
        /// <param name="newValue">The value to which underlyingValue is to be set.</param>
        /// <param name="property">A lambda expression that evaluates to the PropertyInfo of the property 
        /// whose backing-variable reference is supplied.</param>
        /// <returns>True if underlyingValue was updated, otherwise false.</returns>
        protected bool SetProperty<T>(ref T underlyingValue, T newValue, Expression<Func<T>> property)
        {
            bool result = !Object.Equals(newValue, underlyingValue);
            if (result) {
                underlyingValue = newValue;
                OnPropertyChanged(property);
            }
            return result;
        }
    }
}