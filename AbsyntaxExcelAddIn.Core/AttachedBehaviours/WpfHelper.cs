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
using System.Windows;
using System.Windows.Media;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// Static helper class for miscellaneous WPF tasks.
    /// </summary>
    internal static class WpfHelper
    {
        /// <summary>
        /// Finds the first child object of a specific type, optionally with a specific name,
        /// from within the visual tree of a DependencyObject.
        /// </summary>
        public static T FindChild<T>(this DependencyObject obj, string childName) where T : DependencyObject
        {
            T foundChild = null;
            if (obj != null) {
                int childCount = VisualTreeHelper.GetChildrenCount(obj);
                for (int i = 0; i < childCount; i++) {
                    var child = VisualTreeHelper.GetChild(obj, i);
                    if (child.GetType() != typeof(T)) {
                        foundChild = FindChild<T>(child, childName);
                    }
                    else if (!String.IsNullOrEmpty(childName)) {
                        var fe = child as FrameworkElement;
                        if (fe != null && fe.Name == childName) {
                            foundChild = (T)child;
                        }
                        else {
                            foundChild = FindChild<T>(child, childName);
                        }
                    }
                    else {
                        foundChild = (T)child;
                    }
                    if (foundChild != null) break;
                }
            }
            return foundChild;
        }
    }
}