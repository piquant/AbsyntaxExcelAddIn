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
using System.Globalization;
using System.Windows.Data;

namespace AbsyntaxExcelAddIn.Core.Converters
{
    /// <summary>
    /// Converts ProjectExecutionState enum values to <see cref="System.Uri"/> instances that point to PNG 
    /// image files.
    /// </summary>
    [ValueConversion(typeof(ProjectExecutionState), typeof(Uri))]
    public sealed class ProjectExecutionStateToPngUriConverter : PngUriConverterBase
    {
        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ProjectExecutionState result = (ProjectExecutionState)value;
            string fileName;
            switch (result) {
                case ProjectExecutionState.Ineligible:
                    fileName = "Forbidden32";
                    break;
                case ProjectExecutionState.Pending:
                    fileName = "Help32";
                    break;
                case ProjectExecutionState.Executing:
                    fileName = null;
                    break;
                case ProjectExecutionState.Completed:
                    fileName = "Checkmark32";
                    break;
                case ProjectExecutionState.TimedOut:
                case ProjectExecutionState.Aborted:
                case ProjectExecutionState.WriteDataErrors:
                    fileName = "Warning32";    
                    break;
                default:
                    fileName = "Error32";
                    break;
            }
            return Create(fileName);
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}