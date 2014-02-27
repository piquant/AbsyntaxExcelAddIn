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

using System.Collections.Generic;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Defines members that a type implements in order to encapdulate details necessary for loading and 
    /// invoking an Absyntax project.
    /// </summary>
    public interface IProjectInvocationRule
    {
        /// <summary>
        /// Gets a number that identifies the IProjectInvocationRule in a set of such rules.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets or sets a value indicating whether data will be obtained from a worksheet and passed to 
        /// the represented Absyntax project before each invocation.
        /// </summary>
        bool UsesInput { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether any data output by the represented Absyntax project will 
        /// be written to a worksheet after each invocation.
        /// </summary>
        bool UsesOutput { get; set; }

        /// <summary>
        /// Gets or sets the name of the worksheet that will be used to write output data received from the 
        /// represented Absyntax project when UsesOutput is set to true.
        /// </summary>
        string OutputSheetName { get; set; }

        /// <summary>
        /// Gets or sets a notation defining a contiguous range of cells that, when coupled with the selected 
        /// output worksheet, will be used to write data received from the represented Absyntax project after 
        /// each invocation when UsesOutput is set to true.
        /// </summary>
        string OutputCellRange { get; set; }

        /// <summary>
        /// Gets or sets a value which, when combined with the Unit property value, determines the amount of 
        /// time that Absyntax will allow for a project invocation to complete before terminating an 
        /// invocation.
        /// </summary>
        int TimeLimit { get; set; }

        /// <summary>
        /// Gets or sets a value which, when combined with the TimeLimit property value, determines the amount 
        /// of time that Absyntax will allow for a project invocation to complete before terminating an 
        /// invocation.
        /// </summary>
        TimeUnit Unit { get; set; }

        /// <summary>
        /// Gets or sets the full path of the file containing the serialised form of the Absyntax project to
        /// be invoked.
        /// </summary>
        string ProjectPath { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the represented Absyntax project should be reloaded 
        /// prior to each invocation.
        /// </summary>
        bool ReloadProjectBeforeExecuting { get; set; }
        
        /// <summary>
        /// Gets a value indicating whether the IProjectInvocationRule is in a state that allows the Absyntax 
        /// project it represents to be executed.
        /// </summary>
        bool CanExecute { get; }

        /// <summary>
        /// Reads a collection of values from cells in the input range of the selected input data worksheet, 
        /// sequenced in accordance with the input range ordering value.
        /// </summary>
        /// <returns>A collection of cell values.</returns>
        IEnumerable<object> ReadInputData();

        /// <summary>
        /// Attempts to write an object to a range of worksheet cells.
        /// </summary>
        void WriteOutputData(object data);
    }
}