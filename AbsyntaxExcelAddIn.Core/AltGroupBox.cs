using System.Windows.Controls;

namespace AbsyntaxExcelAddIn.Core
{
    /// <summary>
    /// See http://stackoverflow.com/questions/151979/does-the-groupbox-header-in-wpf-swallow-mouse-clicks/3709904#3709904.
    /// </summary>
    public sealed class AltGroupBox : GroupBox
    {
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            var grid = GetVisualChild(0) as Grid;
            if (grid == null || grid.Children.Count <= 3) {
                return;
            }
            var bd = grid.Children[3] as Border;
            if (bd != null) {
                bd.IsHitTestVisible = false;
            }
        }
    }
}