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
using MI2.Events;
using MI2.FrameworkAdapter;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Defines members that a type implements in order to control the runtime load state of Absyntax 
    /// projects.
    /// </summary>
    public interface IRuntimeManager
    {
        /// <summary>
        /// Loads an Absyntax project from file in preparation for invocation.
        /// </summary>
        /// <param name="path">The full path of the file containing the project to be invoked.</param>
        /// <param name="args">Startup arguments used during each invocation.</param>
        /// <returns>A number that uniquely identifies the loaded project.</returns>
        int Load(string path, IStartupArgs args);

        /// <summary>
        /// Synchronously invokes a loaded Absyntax project requiring no input data.
        /// </summary>
        /// <param name="key">The Absyntax runtime's unique identifier of the project to be invoked.</param>
        /// <returns>An object encapsulating the status of the operation and the first output from the 
        /// project.</returns>
        IOperationResult Invoke(int key);

        /// <summary>
        /// Synchronously invokes a loaded Absyntax project requiring input data.
        /// </summary>
        /// <param name="key">The Absyntax runtime's unique identifier of the project to be invoked.</param>
        /// <param name="data">An object to be converted into a type required by the project.</param>
        /// <returns>An object encapsulating the status of the operation and the first output from the 
        /// project.</returns>
        IOperationResult Invoke(int key, object data);

        /// <summary>
        /// Stops and unloads an Absyntax project from the runtime service.
        /// </summary>
        /// <param name="key">A number that uniquely identifies the project to be unloaded.</param>
        void Unload(int key);

        /// <summary>
        /// The event that is raised whenever the hosting service's availability changes.
        /// </summary>
        event EventHandler<EventArgs<bool>> ServiceAvailabilityChanged;
    }
}
