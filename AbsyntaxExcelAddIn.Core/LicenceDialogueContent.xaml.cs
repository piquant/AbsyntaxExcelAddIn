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
using System.Windows;
using System.Windows.Controls;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Interaction logic for LicenceDialogueContent.xaml
    /// </summary>
    public partial class LicenceDialogueContent : UserControl
    {
        public static readonly DependencyProperty UsesFullLicenceProperty =
            DependencyProperty.Register("UsesFullLicence", typeof(bool), typeof(LicenceDialogueContent), new PropertyMetadata(false, OnPropertyChanged));

        public static readonly DependencyProperty ClientIdProperty =
            DependencyProperty.Register("ClientId", typeof(string), typeof(LicenceDialogueContent), new PropertyMetadata(null, OnPropertyChanged));

        private static readonly DependencyPropertyKey IsValidPropertyKey = 
            DependencyProperty.RegisterReadOnly("IsValid", typeof(bool), typeof(LicenceDialogueContent), new PropertyMetadata(false));

        public static readonly DependencyProperty IsValidProperty = IsValidPropertyKey.DependencyProperty;

        private static void OnPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            LicenceDialogueContent c = (LicenceDialogueContent)obj;
            c.UpdateIsValid();
        }

        public LicenceDialogueContent()
        {
            Helper.EnsureApplicationResources();
            InitializeComponent();
            DataContext = this;
        }

        public event EventHandler Accepted;

        public event EventHandler Cancelled;

        private void OnEvent(EventHandler handler)
        {
            if (handler != null) {
                handler(this, EventArgs.Empty);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            OnEvent(Accepted);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            OnEvent(Cancelled);
        }

        private void UpdateIsValid()
        {
            Guid g;
            string value = ClientId;
            bool cidIsValid = Guid.TryParse(value, out g);
            IsValid = UsesFullLicence || cidIsValid;
        }

        public bool UsesFullLicence
        {
            get { return (bool)GetValue(UsesFullLicenceProperty); }
            set { SetValue(UsesFullLicenceProperty, value); }
        }

        public string ClientId
        {
            get { return (string)GetValue(ClientIdProperty); }
            set { SetValue(ClientIdProperty, value); }
        }

        public bool IsValid
        {
            get { return (bool)GetValue(IsValidProperty); }
            private set { SetValue(IsValidPropertyKey, value); }
        }
    }
}