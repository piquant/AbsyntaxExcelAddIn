using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject
{
    [TestClass]
    public class NotifyPropertyChangedBaseTest
    {
        [TestMethod]
        public void RaisesEventWhenPropertyChangedTest()
        {
            var m = new NotifyPropertyChangedBaseMock();
            bool changed = false;
            m.PropertyChanged += (s, e) => changed = true;
            m.Value += 1;
            Assert.IsTrue(changed);
        }

        [TestMethod]
        public void DoesNotRaiseEventWhenPropertyNotChangedTest()
        {
            var m = new NotifyPropertyChangedBaseMock();
            bool changed = false;
            m.PropertyChanged += (s, e) => changed = true;
            int value = m.Value;
            m.Value = value;
            Assert.IsFalse(changed);
        }

        [TestMethod]
        public void EventPropertyNameIsAsExpectedTest()
        {
            var m = new NotifyPropertyChangedBaseMock();
            string propertyName = null;
            m.PropertyChanged += (s, e) => propertyName = e.PropertyName;
            m.Value += 1;
            Assert.AreEqual("Value", propertyName);
        }
    }
}