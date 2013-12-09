using AbsyntaxExcelAddIn.Core;

namespace AbsyntaxExcelAddInTestProject
{
    internal sealed class NotifyPropertyChangedBaseMock : NotifyPropertyChangedBase
    {
        public NotifyPropertyChangedBaseMock()
        { }

        private int m_value;

        public int Value
        {
            get { return m_value; }
            set { SetProperty(ref m_value, value, () => Value); }
        }
    }
}