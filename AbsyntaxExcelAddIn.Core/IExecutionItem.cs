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

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Defines members that a type implements in order provide direct support for the execution of
    /// Absyntax projects.
    /// </summary>
    public interface IExecutionItem
    {
        /// <summary>
        /// Gets the identifier assigned to the project invocation rule represented by the IExecutionItem.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets the identifier assigned by the Absyntax runtime to a project when it was loaded.  A null 
        /// value indicates that the project has not yet been loaded or has not been loaded successfully.
        /// </summary>
        int? Key { get; }

        /// <summary>
        /// Gets the full path to the Absyntax project file to be loaded.
        /// </summary>
        string ProjectPath { get; }

        /// <summary>
        /// Gets the current state of execution.
        /// </summary>
        ProjectExecutionState State { get; }

        /// <summary>
        /// Gets or sets the TextWriter to be used by the Absyntax framework to write project runtime 
        /// messages.
        /// </summary>
        TextWriter Log { get; set; }

        /// <summary>
        /// Starts an asynchronous execution of the represented Absyntax project.
        /// </summary>
        /// <param name="manager">The IRuntimeManager required to perform the underlying project runtime 
        /// tasks.</param>
        /// <param name="callback">The action to be invoked upon completion of the task.</param>
        void BeginExecute(IRuntimeManager manager, Action<IExecutionItem> callback);

        /// <summary>
        /// Aborts a current execution.
        /// </summary>
        /// <param name="manager">The IRuntimeManager required to perform the underlying abort operation.</param>
        void Abort(IRuntimeManager manager);

        /// <summary>
        /// Gets or sets a value indicating whether the IExecutionItem is selected.
        /// </summary>
        bool IsSelected { get; set; }
    }
}