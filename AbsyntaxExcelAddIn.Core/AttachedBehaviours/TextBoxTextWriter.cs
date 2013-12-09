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
using System.Text;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;

namespace AbsyntaxExcelAddIn.Core.AttachedBehaviours
{
    /// <summary>
    /// A TextWriter that writes text to a TextBox.
    /// </summary>
    /// <remarks>
    /// It is conventional to override the Write(char) method, to which all other TextWriter write-methods
    /// defer.  Said method actually does nothing but it is marked as virtual rather than abstract.
    /// <para />
    /// In the context of the TextBoxTextWriter, text is written on threads other than the dispatcher thread 
    /// on which the TextBox was created.  This means that append-text actions must be marshalled to the
    /// dispatcher thread.  It would be excessive to marshal such actions on a per-char basis.  Furthermore,
    /// because we happen to know that the Absyntax framework only ever calls TextWriter.Write and 
    /// TextWriter.WriteLine, this class only overrides these two methods.
    /// <para />
    /// If this class were to be used in more general contexts, the other TextWriter write methods should 
    /// also be overridden.
    /// </remarks>
    internal sealed class TextBoxTextWriter : TextWriter
    {
        /// <summary>
        /// Initialises a new TextBoxTextWriter instances.
        /// </summary>
        /// <param name="output">The TextBoxBase to which text is to be written.</param>
        /// <exception cref="System.ArgumentNullException">output is null.</exception>
        public TextBoxTextWriter(TextBoxBase output)
        {
            Output = output;
        }

        /// <summary>
        /// A weak reference to the wrapped TextBoxBase. allowing this TextBoxTextWriter to be referenced
        /// in contexts other than the user interface without preventing the TextBoxBase from being
        /// garbage-collected.
        private WeakReference m_outputRef;

        /// <summary>
        /// Gets or sets the TextBoxBase to which text is to be written by this TextBoxTextWriter.
        /// </summary>
        public TextBoxBase Output
        {
            get {
                WeakReference wr = m_outputRef;
                return wr == null ? null : wr.Target as TextBoxBase;
            }
            set {
                if (value == null) {
                    throw new ArgumentNullException();
                }
                m_outputRef = new WeakReference(value);
                m_dispatcher = value.Dispatcher;
            }
        }

        /// <summary>
        /// A reference to the wrapped TextBoxBase's Dispatcher.
        /// </summary>
        private Dispatcher m_dispatcher;

        /// <summary>
        /// Writes a string to the wrapped TextBoxBase.
        /// </summary>
        public override void Write(string value)
        {
            base.Write(value);
            Helper.PerformDispatcherAction(m_dispatcher, () => Append(value));
        }

        /// <summary>
        /// Writes a string followed by a line terminator to the wrapped TextBoxBase.
        /// </summary>
        public override void WriteLine(string value)
        {
            base.WriteLine(value);
            Helper.PerformDispatcherAction(m_dispatcher, () => Append(String.Format("{0}{1}", value, Environment.NewLine)));
        }

        /// <summary>
        /// Appends the supplied string to the wrapped TextBoxBase if it is not null.
        /// </summary>
        private void Append(string value)
        {
            TextBoxBase t = Output;
            if (t != null) {
                t.AppendText(value);
            }
        }

        /// <summary>
        /// Returns the <see cref="System.Text.Encoding"/> in which the output is written.
        /// </summary>
        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }
    }
}