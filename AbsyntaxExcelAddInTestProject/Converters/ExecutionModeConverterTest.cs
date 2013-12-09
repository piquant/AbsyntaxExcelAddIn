using System;
using System.Linq;
using AbsyntaxExcelAddIn.Core;
using AbsyntaxExcelAddIn.Core.Converters;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public class ExecutionModeConverterTest
    {
        [TestMethod]
        public void ConvertsEveryEnumValueTest()
        {
            var values = Enum.GetValues(typeof(ExecutionMode)).Cast<ExecutionMode>();
            var c = new ExecutionModeConverter();
            foreach (ExecutionMode value in values) {
                string text = (string)c.Convert(value, null, null, null);
                Assert.IsFalse(String.IsNullOrWhiteSpace(text));
            }
        }
    }
}