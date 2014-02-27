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

using System.IO;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A cache of parameter values used to execute a project via the Absyntax IsolatedRuntimeAdapter.
    /// </summary>
    internal sealed class ProjectExecutionDetail
    {
        /// <summary>
        /// Creates a new ProjectExecutionDetail instance from the details encapsulated in a 
        /// ProjectInvocationRule.
        /// </summary>
        public static ProjectExecutionDetail Create(IProjectInvocationRule rule)
        {
            return new ProjectExecutionDetail() {
                Key = null,
                InputDataRequirement = DataRequirement.None,
                Path = rule.ProjectPath,
                TimeLimit = rule.TimeLimit,
                Unit = rule.Unit,
                Log = null
            };
        }

        /// <summary>
        /// Initialises a new ProjectExecutionDetail instance.
        /// </summary>
        public ProjectExecutionDetail()
        { }

        /// <summary>
        /// Gets or sets a number that uniquely identifies a project loaded via the Absyntax 
        /// IsolatedRuntimeAdapter.  A null value indicates that a project has yet to be loaded.
        /// </summary>
        public int? Key { get; set; }

        /// <summary>
        /// Gets or sets the path of the loaded project.
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Gets or sets the number of time units used to create an execution timeout.
        /// </summary>
        public int TimeLimit { get; set; }

        /// <summary>
        /// Gets or sets the unit of time used to create an execution timeout.
        /// </summary>
        public TimeUnit Unit { get; set; }

        /// <summary>
        /// Gets or sets the project's input data requirement.
        /// </summary>
        public DataRequirement InputDataRequirement { get; set; }

        /// <summary>
        /// Gets or sets the TextWriter to be used to receive messages in respect of an executing project.
        /// </summary>
        public TextWriter Log { get; set; }
    }
}