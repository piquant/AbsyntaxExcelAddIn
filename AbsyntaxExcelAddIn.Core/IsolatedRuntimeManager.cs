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
using MI2.Events;
using MI2.FrameworkAdapter;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Manages fundamental Absyntax Runtime Server process and service tasks.
    /// </summary>
    internal sealed class IsolatedRuntimeManager : IRuntimeManager, IDisposable
    {
        /// <summary>
        /// Initialises a new IsolatedRuntimeManager instance.
        /// </summary>
        public IsolatedRuntimeManager()
        {
            m_adapter = new IsolatedRuntimeAdapter();
        }

        private IsolatedRuntimeAdapter m_adapter;

        /// <summary>
        /// Attempts to open the Absyntax Runtime Server process and start the service that will handle 
        /// Absyntax project execution requests.
        /// </summary>
        /// <param name="clientId">A nullable Guid that, if not null, identifies a third-party execution 
        /// licence permitting project execution.</param>
        /// <exception cref="System.ObjectDisposedException">The IsolatedRuntimeManager instance has been 
        /// disposed.</exception>
        /// <exception cref="System.InvalidOperationException">The process is already open.</exception>
        public void Open(Guid? clientId)
        {
            if (ProcessAvailable) {
                throw new InvalidOperationException("The process is open.  You must close it before reopening it.");
            }
            try {
                if (clientId != null) {
                    m_adapter.SetClientId(clientId.Value);
                }
                m_adapter.Open();
            }
            catch { }
        }

        /// <summary>
        /// Gets a value indicating whether the Absyntax Runtime Server process is open.
        /// </summary>
        /// <remarks>
        /// If the process is unavailable, this can indicate that either the runtime server application
        /// could not be found or that the process has been killed manually.  If the latter is true,
        /// Absyntax will reopen the process straight away.  Once the process is available, Absyntax will 
        /// attempt to make available the project hosting service.
        /// </remarks>
        public bool ProcessAvailable
        {
            get {
                CheckDisposed();
                return m_adapter.ProcessAvailable;
            }
        }

        /// <summary>
        /// The event that is raised whenever the Absyntax Runtime Server process's availability changes.
        /// </summary>
        public event EventHandler<EventArgs<bool>> ProcessAvailabilityChanged
        {
            add {
                CheckDisposed();
                m_adapter.ProcessAvailabilityChanged += value;
            }
            remove {
                CheckDisposed();
                m_adapter.ProcessAvailabilityChanged -= value;
            }
        }

        /// <summary>
        /// The event that is raised whenever the Absyntax Runtime Server process's hosting service
        /// availability changes.
        /// </summary>
        public event EventHandler<EventArgs<bool>> ServiceAvailabilityChanged
        {
            add {
                CheckDisposed();
                m_adapter.ServiceAvailabilityChanged += value;
            }
            remove {
                CheckDisposed();
                m_adapter.ServiceAvailabilityChanged -= value;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the Absyntax project hosting service is available.
        /// </summary>
        /// <exception cref="System.ObjectDisposedException">The IsolatedRuntimeManager instance has been 
        /// disposed.</exception>
        public bool ServiceAvailable
        {
            get {
                CheckDisposed();
                return m_adapter.ServiceAvailable;
            }
        }

        /// <summary>
        /// Loads an Absyntax project from file in preparation for invocation.
        /// </summary>
        /// <param name="path">The full path of the file containing the project to be invoked.</param>
        /// <param name="args">Startup arguments used during each invocation.</param>
        /// <returns>A number that uniquely identifies the loaded project.</returns>
        public int Load(string path, IStartupArgs args)
        {
            CheckDisposed();
            return m_adapter.Load(path, args);
        }

        /// <summary>
        /// Synchronously invokes a loaded Absyntax project requiring no input data.
        /// </summary>
        /// <param name="key">The Absyntax runtime's unique identifier of the project to be invoked.</param>
        /// <returns>An object encapsulating the status of the operation and the first output from the 
        /// project.</returns>
        public IOperationResult Invoke(int key)
        {
            return m_adapter.Invoke(key);
        }

        /// <summary>
        /// Synchronously invokes a loaded Absyntax project requiring input data.
        /// </summary>
        /// <param name="key">The Absyntax runtime's unique identifier of the project to be invoked.</param>
        /// <param name="data">An object to be converted into a type required by the project.</param>
        /// <returns>An object encapsulating the status of the operation and the first output from the 
        /// project.</returns>
        public IOperationResult Invoke(int key, object data)
        {
            return m_adapter.Invoke(key, data);
        }

        /// <summary>
        /// Stops and unloads an Absyntax project from the runtime service.
        /// </summary>
        /// <param name="key">A number that uniquely identifies the project to be unloaded.</param>
        public void Unload(int key)
        {
            CheckDisposed();
            try {
                m_adapter.Unload(key);
            }
            catch (InvalidOperationException) { }
        }

        /// <summary>
        /// Closes the Absyntax Runtime Server process if it is open.
        /// </summary>
        public void Close()
        {
            if (ProcessAvailable) {
                m_adapter.Close();
            }
        }

        /// <summary>
        /// Disposes this IsolatedRuntimeManager.
        /// </summary>
        public void Dispose()
        {
            var adapter = m_adapter;
            if (adapter != null) {
                adapter.Dispose();
                m_adapter = null;
            }
        }

        /// <summary>
        /// Throws a <see cref="System.ObjectDisposedException"/> if this IsolatedRuntimeManager has been 
        /// disposed.
        /// </summary>
        private void CheckDisposed()
        {
            if (m_adapter == null) {
                throw new ObjectDisposedException(typeof(IsolatedRuntimeManager).Name);
            }
        }
    }
}