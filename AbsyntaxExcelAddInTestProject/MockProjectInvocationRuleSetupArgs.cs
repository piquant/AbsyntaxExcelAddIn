using System;
using System.Collections.Generic;
using AbsyntaxExcelAddIn.Core;

namespace AbsyntaxExcelAddInTestProject
{
    internal sealed class MockProjectInvocationRuleSetupArgs
    {
        public int Id { get; set; }

        public bool UsesInput { get; set; }

        public bool UsesOutput { get; set; }

        public string OutputSheetName { get; set; }

        public string OutputCellRange { get; set; }

        public int TimeLimit { get; set; }

        public TimeUnit Unit { get; set; }

        public string ProjectPath { get; set; }

        public bool CanExecute { get; set; }

        public IEnumerable<object> ReadInputData { get; set; }

        public Action<object> WriteOutputData { get; set; }
    }
}