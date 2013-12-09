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
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// Interaction logic for InProgressIndicator.xaml
    /// </summary>
    /// <remarks>
    /// Adapted from http://sachabarbs.wordpress.com/2009/12/29/better-wpf-circular-progress-bar/.
    /// </remarks>
    public partial class InProgressIndicator : UserControl
    {
        private static readonly DependencyPropertyKey EllipseSizePropertyKey =
            DependencyProperty.RegisterReadOnly("EllipseSize", typeof(double), typeof(InProgressIndicator), new PropertyMetadata(OnEllipseSizeChanged));

        public static readonly DependencyProperty EllipseSizeProperty = EllipseSizePropertyKey.DependencyProperty;

        private static void OnEllipseSizeChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            InProgressIndicator ipi = obj as InProgressIndicator;
            ipi.SetPositions();
        }

        private readonly DispatcherTimer m_animationTimer;
        
        public InProgressIndicator()
        {
            InitializeComponent();
            m_animationTimer = new DispatcherTimer(DispatcherPriority.ContextIdle, Dispatcher);
            m_animationTimer.Interval = new TimeSpan(0, 0, 0, 0, 100);
        }

        public double EllipseSize
        {
            get { return (double)GetValue(EllipseSizeProperty); }
            private set { SetValue(EllipseSizePropertyKey, value); }
        }

        private void Start()
        {
            m_animationTimer.Tick += AnimationTimer_Tick;
            m_animationTimer.Start();
        }

        private void Stop()
        {
            m_animationTimer.Stop();
            m_animationTimer.Tick -= AnimationTimer_Tick;
        }

        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            double i = C0.Opacity;
            C0.Opacity = C1.Opacity;
            C1.Opacity = C2.Opacity;
            C2.Opacity = C3.Opacity;
            C3.Opacity = C4.Opacity;
            C4.Opacity = C5.Opacity;
            C5.Opacity = C6.Opacity;
            C6.Opacity = C7.Opacity;
            C7.Opacity = C8.Opacity;
            C8.Opacity = C9.Opacity;
            C9.Opacity = i;
        }

        private void SetPositions()
        {
            double widthSpace = (ActualWidth - EllipseSize) / 2.0;
            double heightSpace = (ActualHeight - EllipseSize) / 2.0;
            double scale = Math.Min(widthSpace, heightSpace);
            SetPosition(C0, 0.0, scale);
            SetPosition(C1, 1.0, scale);
            SetPosition(C2, 2.0, scale);
            SetPosition(C3, 3.0, scale);
            SetPosition(C4, 4.0, scale);
            SetPosition(C5, 5.0, scale);
            SetPosition(C6, 6.0, scale);
            SetPosition(C7, 7.0, scale);
            SetPosition(C8, 8.0, scale);
            SetPosition(C9, 9.0, scale);
        }

        private void SetPosition(Ellipse ellipse, double index, double scale)
        {
            const double step = Math.PI * 2 / 10.0;
            ellipse.SetValue(Canvas.LeftProperty, Math.Sin(index * step) * scale);
            ellipse.SetValue(Canvas.TopProperty, Math.Cos(index * step) * scale);
        }

        private void Canvas_Unloaded(object sender, RoutedEventArgs e)
        {
            Stop();
        }

        private void UserControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue) {
                Start();
            }
            else {
                Stop();
            }
        }

        private void UserControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            EllipseSize = Math.Min(ActualWidth, ActualHeight) / 6.0;
        }
    }
}