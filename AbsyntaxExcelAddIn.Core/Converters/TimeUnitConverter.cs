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
using System.Globalization;
using System.Windows.Data;
using AbsyntaxExcelAddIn.Resources;

namespace AbsyntaxExcelAddIn.Core.Converters
{
    /// <summary>
    /// Converts TimeUnit enum values to and from strings.
    /// </summary>
    [ValueConversion(typeof(TimeUnit), typeof(string))]
    public sealed class TimeUnitConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            TimeUnit tu = (TimeUnit)value;
            switch (tu) {
                case TimeUnit.Seconds:
                    return TextResources.Seconds;
                case TimeUnit.Minutes:
                    return TextResources.Minutes;
                case TimeUnit.Hours:
                    return TextResources.Hours;
                default:
                    return TextResources.Days;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string s = (string)value;
            TimeUnit tu;
            if (s == TextResources.Seconds) {
                tu = TimeUnit.Seconds;
            }
            else if (s == TextResources.Minutes) {
                tu = TimeUnit.Minutes;
            }
            else if (s == TextResources.Hours) {
                tu = TimeUnit.Hours;
            }
            else {
                tu = TimeUnit.Days;
            }
            return tu;
        }
    }
}