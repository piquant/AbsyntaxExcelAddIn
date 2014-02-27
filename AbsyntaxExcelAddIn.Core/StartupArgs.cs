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
using System.IO;
using MI2.FrameworkAdapter;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A simple implementation of the <see cref="MI2.FrameworkAdapter.IStartupArgs"/> interface.
    /// </summary>
    internal sealed class StartupArgs : IStartupArgs
    {
        /// <summary>
        /// Initialises a new StartupArgs instance.
        /// </summary>
        public StartupArgs()
        { }

        /// <summary>
        /// Gets or sets a TextWriter to which log messages from the Absyntax runtime host can be written.
        /// </summary>
        public TextWriter Log { get; set; }

        /// <summary>
        /// Gets or sets the timeout period (in milliseconds) within which each operation is to be completed.
        /// </summary>
        public int OperationTimeout { get; set; }

        /// <summary>
        /// Gets a value indicating whether each operation is to be completed within a timeout period.  A 
        /// <see cref="System.NotSupportedException"/> is thrown if an attempt is made to set this value.  
        /// This is because the add-in enforces the use of runtime timeouts.
        /// </summary>
        public bool UseOperationTimeout
        {
            get { return true; }
            set { throw new NotSupportedException(); }
        }
    }
}