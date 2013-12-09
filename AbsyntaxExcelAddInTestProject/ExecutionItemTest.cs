using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using AbsyntaxExcelAddIn.Core;
using AbsyntaxExcelAddIn.Resources;
using MI2.FrameworkAdapter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace AbsyntaxExcelAddInTestProject
{
    [TestClass]
    public class ExecutionItemTest
    {
        private static IProjectInvocationRule CreateRule(MockProjectInvocationRuleSetupArgs args)
        {
            var mock = new Mock<IProjectInvocationRule>();
            if (args != null) {
                mock.Setup(m => m.Id).Returns(args.Id);
                mock.Setup(m => m.UsesInput).Returns(args.UsesInput);
                mock.Setup(m => m.UsesOutput).Returns(args.UsesOutput);
                mock.Setup(m => m.OutputSheetName).Returns(args.OutputSheetName);
                mock.Setup(m => m.OutputCellRange).Returns(args.OutputCellRange);
                mock.Setup(m => m.TimeLimit).Returns(args.TimeLimit);
                mock.Setup(m => m.Unit).Returns(args.Unit);
                mock.Setup(m => m.ProjectPath).Returns(args.ProjectPath);
                mock.Setup(m => m.CanExecute).Returns(args.CanExecute);
                mock.Setup(m => m.ReadInputData()).Returns(args.ReadInputData);
                Action<object> a = args.WriteOutputData;
                if (a != null) {
                    mock.Setup(m => m.WriteOutputData(It.IsAny<object>())).Callback(a);
                }
            }
            return mock.Object;
        }

        private static ExecutionItem CreateItem(MockProjectInvocationRuleSetupArgs args = null, ProjectExecutionDetail detail = null)
        {
            IProjectInvocationRule r = CreateRule(args);
            ProjectExecutionDetail d = detail ?? ProjectExecutionDetail.Create(r);
            return new ExecutionItem(r, d);
        }

        [TestMethod]
        public void IsIneligibleForExecutionIfRuleCannotExecuteTest()
        {
            PerformProjectExecutionStateTest(false, ProjectExecutionState.Ineligible);
        }

        [TestMethod]
        public void IsEligibleForExecutionIfRuleCanExecuteTest()
        {
            PerformProjectExecutionStateTest(true, ProjectExecutionState.Pending);
        }

        private void PerformProjectExecutionStateTest(bool canExecute, ProjectExecutionState expectedState)
        {
            var args = new MockProjectInvocationRuleSetupArgs() { CanExecute = canExecute };
            var i = CreateItem(args);
            Assert.AreEqual(expectedState, i.State);
        }

        [TestMethod]
        public void AutoCreateLogIsTrueIfExecutionDetailLogIsNullTest()
        {
            PerformAutoCreateLogTest(null, true);
        }

        [TestMethod]
        public void AutoCreateLogIsFalseIfExecutionDetailLogIsNotNullTest()
        {
            PerformAutoCreateLogTest(new Mock<TextWriter>().Object, false);
        }

        private void PerformAutoCreateLogTest(TextWriter log, bool expectedResult)
        {
            var d = new ProjectExecutionDetail() { Log = log };
            var i = CreateItem(null, d);
            Assert.AreEqual(expectedResult, i.AutoCreateLog);
        }

        [TestMethod]
        public void ProjectPathReturnsPlaceholderTextIfRulePathIsNullTest()
        {
            PerformProjectPathTest(null, TextResources.ProjectPathNotSpecified);
        }

        [TestMethod]
        public void ProjectPathReturnsPlaceholderTextIfRulePathIsEmptyTest()
        {
            PerformProjectPathTest(String.Empty, TextResources.ProjectPathNotSpecified);
        }

        [TestMethod]
        public void ProjectPathReturnsPlaceholderTextIfRulePathIsWhitespaceTest()
        {
            PerformProjectPathTest(" ", TextResources.ProjectPathNotSpecified);
        }

        [TestMethod]
        public void ProjectPathReturnsRulePathIfRulePathIsNotEmptyOrWhitespaceTest()
        {
            string path = "path";
            PerformProjectPathTest(path, path);
        }

        private void PerformProjectPathTest(string ruleProjectPath, string expectedItemProjectPath)
        {
            var args = new MockProjectInvocationRuleSetupArgs() { ProjectPath = ruleProjectPath };
            var i = CreateItem(args);
            Assert.AreEqual(expectedItemProjectPath, i.ProjectPath);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void BeginExecuteThrowsExceptionIfStateIsNotPendingTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() { CanExecute = false };
            var i = CreateItem(args);
            var rm = new Mock<IRuntimeManager>().Object;
            i.BeginExecute(rm, ei => { });
        }

        [TestMethod]
        public void BeginExecuteLoadsProjectIfNotAlreadyLoadedTest()
        {
            PerformLoadVerificationTest(null, Times.Once()); // Null key denotes non-loaded project
        }

        [TestMethod]
        public void BeginExecuteDoesNotLoadProjectIfAlreadyLoadedTest()
        {
            PerformLoadVerificationTest(1, Times.Never()); // Non-null key denotes loaded project
        }

        private void PerformLoadVerificationTest(int? detailKey, Times loadInvocationTimes)
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true, 
                ProjectPath = "path", 
                UsesInput = false
            };
            var d = new ProjectExecutionDetail() { Key = detailKey };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            rmMock.Verify(m => m.Load(args.ProjectPath, It.IsAny<IStartupArgs>()), loadInvocationTimes);
        }
        
        [TestMethod]
        public void BeginExecuteInvokesProjectWithNoDataIfUsesInputIsFalseTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true, 
                UsesInput = false
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            rmMock.Verify(m => m.Invoke(key), Times.Once());
            rmMock.Verify(m => m.Invoke(key, It.IsAny<IEnumerable<object>>()), Times.Never());
        }

        [TestMethod]
        public void BeginExecuteInvokesProjectWithDataIfUsesInputIsTrueTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true, 
                UsesInput = true,
                ReadInputData = new object[1]
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            rmMock.Verify(m => m.Invoke(key), Times.Never());
            rmMock.Verify(m => m.Invoke(key, It.IsAny<IEnumerable<object>>()), Times.Once());
        }

        [TestMethod]
        public void IsExecutingReturnsTrueWhileExecutingTest()
        {
            ExecutionItem i;
            IRuntimeManager manager;
            SetupForLongInvocation(out i, out manager);
            var mre = new ManualResetEvent(false);
            Assert.IsFalse(i.IsExecuting);
            i.BeginExecute(manager, ei => mre.Set());
            Thread.Sleep(100);
            Assert.IsTrue(i.IsExecuting);
            mre.WaitOne(200);
            Assert.IsFalse(i.IsExecuting);
        }

        [TestMethod]
        public void StateIsExecutingWhileExecutingTest()
        {
            ExecutionItem i;
            IRuntimeManager manager;
            SetupForLongInvocation(out i, out manager);
            var mre = new ManualResetEvent(false);
            Assert.AreNotEqual(ProjectExecutionState.Executing, i.State);
            i.BeginExecute(manager, ei => mre.Set());
            Thread.Sleep(100);
            Assert.AreEqual(ProjectExecutionState.Executing, i.State);
            mre.WaitOne(200);
            Assert.AreNotEqual(ProjectExecutionState.Executing, i.State);
        }

        private static void SetupForLongInvocation(out ExecutionItem i, out IRuntimeManager manager)
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            rmMock.Setup(m => m.Invoke(key)).Callback(() => Thread.Sleep(200));
            manager = rmMock.Object;
        }

        [TestMethod]
        public void NullReturnFromInvocationIndicatesAbortedProjectTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            rmMock.Setup(m => m.Invoke(key)).Returns((IOperationResult)null);
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            Assert.AreEqual(ProjectExecutionState.Aborted, i.State);
        }

        /// <summary>
        /// If the status of a project invocation is anything other than OperationStatus.Ok, the project
        /// will need reloading.  This is indicated by setting ExecutionItem.Key to null.
        /// </summary>
        [TestMethod]
        public void KeyIsSetToNullIfInvokedOperationStatusIsNotOkTest()
        {
            var notOkStatuses = Enum.GetValues(typeof(OperationStatus)).Cast<OperationStatus>().Except(new OperationStatus[] { OperationStatus.Ok });
            foreach (OperationStatus s in notOkStatuses) {
                PerformKeyTest(s, true);
            }
        }

        [TestMethod]
        public void KeyIsNotSetToNullIfInvokedOperationStatusIsOkTest()
        {
            PerformKeyTest(OperationStatus.Ok, false);
        }

        private void PerformKeyTest(OperationStatus result, bool expectedKeyIsNull)
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var orMock = new Mock<IOperationResult>();
            orMock.Setup(m => m.Status).Returns(result);
            rmMock.Setup(m => m.Invoke(key)).Returns(orMock.Object);
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            if (expectedKeyIsNull) {
                Assert.IsNull(i.Key);
            }
            else {
                Assert.IsNotNull(i.Key);
                Assert.AreEqual(key, i.Key.Value);
            }
        }

        [TestMethod]
        public void WritesOutputDataIfInvocationCompletedSuccessfullyAndRuleExpectsOutputTest()
        {
            PerformWritesOutputDataTest(OperationStatus.Ok, true);
        }

        [TestMethod]
        public void DoesNotWriteOutputDataIfInvocationCompletedSuccessfullyAndRuleDoesNotExpectOutputTest()
        {
            PerformWritesOutputDataTest(OperationStatus.Ok, false);
        }

        [TestMethod]
        public void DoesNotWriteOutputDataIfInvocationDoesNotCompleteSuccessfullyTest()
        {
            PerformWritesOutputDataTest(OperationStatus.ServiceEndUnknown, true);
        }

        private void PerformWritesOutputDataTest(OperationStatus result, bool usesOutput)
        {
            object writtenData = null;
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false,
                UsesOutput = usesOutput,
                WriteOutputData = new Action<object>(o => writtenData = o)
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            object data = new object();
            var orMock = new Mock<IOperationResult>();
            orMock.Setup(m => m.Status).Returns(result);
            orMock.Setup(m => m.Data).Returns(data);
            rmMock.Setup(m => m.Invoke(key)).Returns(orMock.Object);
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            if (result == OperationStatus.Ok && usesOutput) {
                Assert.AreSame(data, writtenData);
            }
            else {
                Assert.IsNull(writtenData);
            }
        }

        [TestMethod]
        public void InvocationErrorMessageWrittenToLogTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false,
            };
            int key = 1;
            var logMock = new Mock<TextWriter>();
            var d = new ProjectExecutionDetail() { Key = key, Log = logMock.Object };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            string errMsg = "This is an error message";
            rmMock.Setup(m => m.Invoke(key)).Callback(() => { throw new Exception(errMsg); });
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            logMock.Verify(m => m.WriteLine(errMsg));
        }

        [TestMethod]
        public void CallbackPassesReferenceToExecutionItemTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false,
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rm = new Mock<IRuntimeManager>().Object;
            var mre = new ManualResetEvent(false);
            IExecutionItem callbackArg = null;
            i.BeginExecute(rm, ei => {
                callbackArg = ei;
                mre.Set();
            });
            mre.WaitOne(200);
            Assert.AreSame(i, callbackArg);
        }

        [TestMethod]
        public void ProjectReloadedIfDataConversionProblemTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                ProjectPath = "path", 
                UsesInput = true,
                ReadInputData = new object[1]
            };
            var d = new ProjectExecutionDetail() { Key = 1 };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var orMock = new Mock<IOperationResult>();
            orMock.Setup(m => m.Status).Returns(OperationStatus.StartupDataConversionProblem);
            rmMock.Setup(m => m.Invoke(It.IsAny<int>(), It.IsAny<object>())).Returns(orMock.Object);
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            rmMock.Verify(m => m.Load(args.ProjectPath, It.IsAny<IStartupArgs>()), Times.AtLeastOnce());
        }

        /// <summary>
        /// Presenting an enumerable containing one non-null value should result in three attempts to 
        /// invoke a project if each invocation keeps returning OperationStatus.StartupDataConversionProblem.
        /// </summary>
        [TestMethod]
        public void UpToThreeDataAggregationsAttemptedTest()
        {
            int[] data = { 46 };
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                ProjectPath = "path", 
                UsesInput = true,
                ReadInputData = data.Cast<object>()
            };
            var d = new ProjectExecutionDetail() { Key = 1 };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            var orMock = new Mock<IOperationResult>();
            orMock.Setup(m => m.Status).Returns(OperationStatus.StartupDataConversionProblem);
            rmMock.Setup(m => m.Invoke(It.IsAny<int>(), It.IsAny<object>())).Returns(orMock.Object);
            var mre = new ManualResetEvent(false);
            i.BeginExecute(rmMock.Object, ei => mre.Set());
            mre.WaitOne(200);
            rmMock.Verify(m => m.Invoke(It.IsAny<int>(), It.IsAny<object>()), Times.Exactly(3));
        }

        [TestMethod]
        public void UnloadAttemptNotMadeIfAbortCalledWhileNotExecutingTest()
        {
            var i = CreateItem();
            var rmMock = new Mock<IRuntimeManager>();
            i.Abort(rmMock.Object);
            rmMock.Verify(m => m.Unload(It.IsAny<int>()), Times.Never());
        }

        [TestMethod]
        public void UnloadAttemptedIfAbortCalledWhileExecutingTest()
        {
            var args = new MockProjectInvocationRuleSetupArgs() {
                CanExecute = true,
                UsesInput = false
            };
            int key = 1;
            var d = new ProjectExecutionDetail() { Key = key };
            var i = CreateItem(args, d);
            var rmMock = new Mock<IRuntimeManager>();
            rmMock.Setup(m => m.Invoke(key)).Callback(() => Thread.Sleep(200));
            i.BeginExecute(rmMock.Object, null);
            Thread.Sleep(100);
            i.Abort(rmMock.Object);
            rmMock.Verify(m => m.Unload(key), Times.Once());
        }
    }
}