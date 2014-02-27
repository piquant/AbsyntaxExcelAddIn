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
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.IO.IsolatedStorage;
using System.Runtime.Serialization.Formatters.Binary;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// A settings provider that caches its settings in an isolated storage file scoped by assembly and 
    /// roaming user identities.
    /// </summary>
    public sealed class IsolatedStorageSettingsProvider : SettingsProvider
    {
        /// <summary>
        /// Initialises a new IsolatedStorageSettingsProvider instance.
        /// </summary>
        public IsolatedStorageSettingsProvider()
        { }

        public override string Name
        {
            get { return typeof(IsolatedStorageSettingsProvider).Name; }
        }

        /// <summary>
        /// Gets or sets the name of the currently running application.
        /// </summary>
        public override string ApplicationName
        {
            get { return "AbsyntaxExcelAddIn"; }
            set { }
        }

        public override void Initialize(string name, NameValueCollection col)
        {
            base.Initialize(this.ApplicationName, col);
        }

        private static readonly string s_fileName = "settings.config";

        /// <summary>
        /// Returns a collection of settings property values.
        /// </summary>
        public override SettingsPropertyValueCollection GetPropertyValues(SettingsContext context, SettingsPropertyCollection collection)
        {
            var spvc = new SettingsPropertyValueCollection();
            foreach (SettingsProperty sp in collection) {
                SettingsPropertyValue value = new SettingsPropertyValue(sp);
                value.IsDirty = false;
                value.SerializedValue = GetSettingValue(sp);
                spvc.Add(value);
            }
            return spvc;
        }

        private object GetSettingValue(SettingsProperty setting)
        {
            var cache = GetCache();
            object value;
            cache.TryGetValue(setting.Name, out value);
            return value;
        }

        /// <summary>
        /// Persists a collection of property values.
        /// </summary>
        public override void SetPropertyValues(SettingsContext context, SettingsPropertyValueCollection propvals)
        {
            foreach (SettingsPropertyValue propval in propvals) {
                SetSettingValue(propval.Name, propval.SerializedValue);
            }
            PersistCache();
        }

        private void SetSettingValue(string name, object value)
        {
            var cache = GetCache();
            cache[name] = value;
        }

        private void PersistCache()
        {
            var cache = GetCache();
            var formatter = new BinaryFormatter();
            lock (m_lockObj) {
                using (var store = GetStore())
                using (var stream = store.OpenFile(s_fileName, FileMode.OpenOrCreate, FileAccess.Write)) {
                    formatter.Serialize(stream, cache);
                }
            }
        }

        private object m_lockObj = new object();

        private volatile Dictionary<string, object> m_cache;

        private Dictionary<string, object> GetCache()
        {
            if (m_cache == null) {
                lock (m_lockObj) {
                    if (m_cache == null) {
                        m_cache = CreateCache();
                    }
                }
            }
            return m_cache;
        }

        private Dictionary<string, object> CreateCache()
        {
            Dictionary<string, object> cache = null;
            lock (m_lockObj) {
                using (IsolatedStorageFile store = GetStore())
                using (IsolatedStorageFileStream stream = store.OpenFile(s_fileName, FileMode.OpenOrCreate, FileAccess.Read)) {
                    if (stream.Length > 0) {
                        var formatter = new BinaryFormatter();
                        try {
                            cache = (Dictionary<string, object>)formatter.Deserialize(stream);
                        }
                        catch { }
                    }
                }
            }
            return cache ?? new Dictionary<string, object>();
        }

        private IsolatedStorageFile GetStore()
        {
            return IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly | IsolatedStorageScope.Roaming, null, null);
        }
    }
}