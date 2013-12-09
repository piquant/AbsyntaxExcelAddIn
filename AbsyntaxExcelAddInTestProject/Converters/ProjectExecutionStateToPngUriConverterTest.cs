using System.Collections.Generic;
using System.Windows.Data;
using AbsyntaxExcelAddIn.Core;
using AbsyntaxExcelAddIn.Core.Converters;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public class ProjectExecutionStateToPngUriConverterTest : PngUriConverterBaseTest<ProjectExecutionState>
    {
        protected override IValueConverter GetConverter()
        {
            return new ProjectExecutionStateToPngUriConverter();
        }

        protected override IEnumerable<ProjectExecutionState> Exclusions
        {
            get { return new ProjectExecutionState[] { ProjectExecutionState.Executing }; }
        }
    }
}