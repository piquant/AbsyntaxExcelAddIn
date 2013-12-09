/* Copyright (c) 2007-2010, Adolfo Marinucci
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without modification, are permitted 
 * provided that the following conditions are met:
 * 
 * Redistributions of source code must retain the above copyright notice, this list of conditions 
 * and the following disclaimer.
 * 
 * Redistributions in binary form must reproduce the above copyright notice, this list of conditions 
 * and the following disclaimer in the documentation and/or other materials provided with the distribution.
 * 
 * Neither the name of Adolfo Marinucci nor the names of its contributors may be used to endorse or 
 * promote products derived from this software without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED 
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A 
 * PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR 
 * ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT 
 * LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR 
 * TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * */

using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// An Image that becomes semi-opaque when disabled.
    /// </summary>
    public class AutoGreyableImage : Image
    {
        private Brush m_opacityMaskC, m_opacityMaskG;
        
        private ImageSource m_sourceC, m_sourceG;

        static AutoGreyableImage()
        {
            FrameworkElement.DefaultStyleKeyProperty.OverrideMetadata(typeof(AutoGreyableImage), new FrameworkPropertyMetadata(typeof(AutoGreyableImage)));
        }

        /// <summary>
        /// Initialises a new AutoGreyableImage instance.
        /// </summary>
        public AutoGreyableImage()
        { }

        protected override void OnPropertyChanged(DependencyPropertyChangedEventArgs e)
        {
            string propertyName = e.Property.Name;
            if (propertyName == "IsEnabled") {
                object newValue = e.NewValue;
                if ((newValue as bool?) == false) {
                    Source = m_sourceG;
                    OpacityMask = m_opacityMaskG;
                }
                else if ((newValue as bool?) == true) {
                    Source = m_sourceC;
                    OpacityMask = m_opacityMaskC;
                }
            }
            else if ((propertyName == "Source" && !Object.ReferenceEquals(Source, m_sourceC)) && !Object.ReferenceEquals(Source, m_sourceG)) {
                SetSources();
            }
            else if ((propertyName == "OpacityMask" && !Object.ReferenceEquals(OpacityMask, m_opacityMaskC)) && !Object.ReferenceEquals(OpacityMask, m_opacityMaskG)) {
                m_opacityMaskC = OpacityMask;
            }
            base.OnPropertyChanged(e);
        }

        private void SetSources()
        {
            m_sourceG = m_sourceC = Source;
            m_opacityMaskG = new ImageBrush(m_sourceC);
            m_opacityMaskG.Opacity = 1;
            try {
                string stringUri = TypeDescriptor.GetConverter(Source).ConvertTo(Source, typeof(string)) as string;
                Uri uri = null;
                if (!Uri.TryCreate(stringUri, UriKind.Absolute, out uri)) {
                    uri = new Uri("pack://application:,,,/" + stringUri.TrimStart(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }));
                }
                m_sourceG = new FormatConvertedBitmap(new BitmapImage(uri), PixelFormats.Pbgra32, null, 0.0);
            }
            catch { }
        }
    }
}