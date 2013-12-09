﻿/* Copyright © 2013 Managing Infrastructure Information Ltd
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

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Defines the range of pre- and post-execution states applicable to an Absyntax project.
    /// </summary>
    public enum ProjectExecutionState
    {
        /// <summary>
        /// Indicates that a project will not be executed.
        /// </summary>
        Ineligible,

        /// <summary>
        /// Indicates that a project is awaiting execution.
        /// </summary>
        Pending,

        /// <summary>
        /// Indicates that a project is in the process of being executed.
        /// </summary>
        Executing,

        /// <summary>
        /// Indicates that a project was executed successfully.
        /// </summary>
        Completed,

        /// <summary>
        /// Indicates that a project was executed successfully but not all of its output data could be 
        /// written to the target worksheet.
        /// </summary>
        WriteDataErrors,

        /// <summary>
        /// Indicates that a project was executed but timed out before it could complete its tasks.
        /// </summary>
        TimedOut,

        /// <summary>
        /// Indicates that a project has been aborted during execution.
        /// </summary>
        Aborted,

        /// <summary>
        /// Indicates that an attempt was made to execute a project but errors occurred during the 
        /// operation.
        /// </summary>
        Errors
    }
}