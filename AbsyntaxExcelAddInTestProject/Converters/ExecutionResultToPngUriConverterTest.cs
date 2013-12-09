using System.Windows.Data;
using AbsyntaxExcelAddIn.Core;
using AbsyntaxExcelAddIn.Core.Converters;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public class ExecutionResultToPngUriConverterTest : PngUriConverterBaseTest<ExecutionResult>
    {
        protected override IValueConverter GetConverter()
        {
            return new ExecutionResultToPngUriConverter();
        }
    }
}