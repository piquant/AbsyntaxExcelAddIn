using AbsyntaxExcelAddIn.Core.Converters;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public class BoolConverterBaseTest
    {
        [TestMethod]
        public void ConvertsBooleanValuesAsExpectedTest()
        {
            var cMock = new Mock<BoolConverterBase<int>>();
            cMock.CallBase = true;
            var c = cMock.Object;
            int trueValue = 1, falseValue = -1;
            c.TrueValue = trueValue;
            c.FalseValue = falseValue;
            object v = c.Convert(true, null, null, null);
            Assert.AreEqual(trueValue, v);
            v = c.Convert(false, null, null, null);
            Assert.AreEqual(falseValue, v);
        }

        [TestMethod]
        public void ConvertsConvertedValuesBackAsExpectedTest()
        {
            var cMock = new Mock<BoolConverterBase<int>>();
            cMock.CallBase = true;
            var c = cMock.Object;
            int trueValue = 1, falseValue = -1;
            c.TrueValue = trueValue;
            c.FalseValue = falseValue;
            object v = c.ConvertBack(trueValue, null, null, null);
            Assert.AreEqual(true, v);
            v = c.ConvertBack(falseValue, null, null, null);
            Assert.AreEqual(false, v);
            v = c.ConvertBack(0, null, null, null);
            Assert.AreNotEqual(true, v);
            Assert.AreNotEqual(false, v);
        }
    }
}