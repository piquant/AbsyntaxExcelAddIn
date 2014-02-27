using System;
using System.IO;
using System.Reflection;
using AbsyntaxExcelAddIn.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace AbsyntaxExcelAddInTestProject
{
    /// <summary>
    /// There are insurmountable issues mocking Excel interop interfaces, meaning that there are
    /// far fewer tests herein that there should be.
    /// </summary>
    [TestClass]
    public class ProjectInvocationRuleTest
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            var a = Assembly.GetExecutingAssembly();
            string s = a.Location;
            string ap = Path.GetDirectoryName(s);
            s_tempFilePath = Path.Combine(ap, "Temp.apj");
            using (var fs = File.Create(s_tempFilePath)) { }
        }

        private static string s_tempFilePath;

        [ClassCleanup]
        public static void ClassCleanup()
        {
            File.Delete(s_tempFilePath);
        }

        private static ProjectInvocationRule Create(
            IWorksheetProvider wsProvider = null, 
            INamedRangeProvider nrProvider = null, 
            IDataReader reader = null)
        {
            wsProvider = wsProvider ?? new Mock<IWorksheetProvider>().Object;
            nrProvider = nrProvider ?? new Mock<INamedRangeProvider>().Object;
            reader = reader ?? new Mock<IDataReader>().Object;
            return new ProjectInvocationRule(wsProvider, nrProvider, reader);
        }

        private static ProjectInvocationRule Create(
            int id,
            IWorksheetProvider wsProvider = null,
            INamedRangeProvider nrProvider = null)
        {
            wsProvider = wsProvider ?? new Mock<IWorksheetProvider>().Object;
            nrProvider = nrProvider ?? new Mock<INamedRangeProvider>().Object;
            return new ProjectInvocationRule(wsProvider, nrProvider, id);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullWorksheetProviderThrowsException1Test()
        {
            var nrp = new Mock<INamedRangeProvider>().Object;
            new ProjectInvocationRule(null, nrp, 1);
        }
        
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullWorksheetProviderThrowsException2Test()
        {
            var nrp = new Mock<INamedRangeProvider>().Object;
            var dr = new Mock<IDataReader>().Object;
            new ProjectInvocationRule(null, nrp, dr);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullNamedRangeProviderThrowsException1Test()
        {
            var wsp = new Mock<IWorksheetProvider>().Object;
            new ProjectInvocationRule(wsp, null, 1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullNamedRangeProviderThrowsException2Test()
        {
            var wsp = new Mock<IWorksheetProvider>().Object;
            var dr = new Mock<IDataReader>().Object;
            new ProjectInvocationRule(wsp, null, dr);
        }

        [TestMethod]
        public void DeserialisationConstructorTest()
        {
            var drMock = new Mock<IDataReader>();
            var r = Create(null, null, drMock.Object);
            drMock.Verify(m => m.Read<int>(), Times.AtLeastOnce());
            drMock.Verify(m => m.Read<bool>(), Times.AtLeastOnce());
            drMock.Verify(m => m.Read<string>(), Times.AtLeastOnce());
            drMock.Verify(m => m.Read<RangeOrdering>(), Times.AtLeastOnce());
            drMock.Verify(m => m.Read<TimeUnit>(), Times.AtLeastOnce());
            drMock.Verify(m => m.Read<ExecutionResult>(), Times.AtLeastOnce());
        }

        [TestMethod]
        public void InputSheetKeyIsNullIfNoSheetsToSelectTest()
        {
            var r = Create(1);
            Assert.IsNull(r.InputSheetKey);
        }

        [TestMethod]
        public void InputSheetNameIsNullIfNoSheetsToSelectTest()
        {
            var r = Create(1);
            Assert.IsNull(r.InputSheetName);
        }

        [TestMethod]
        public void OutputSheetKeyIsNullIfNoSheetsToSelectTest()
        {
            var r = Create(1);
            Assert.IsNull(r.OutputSheetKey);
        }

        [TestMethod]
        public void OutputSheetNameIsNullIfNoSheetsToSelectTest()
        {
            var r = Create(1);
            Assert.IsNull(r.OutputSheetName);
        }

        [TestMethod]
        public void IsNotValidIfProjectPathIsNullTest()
        {
            PerformProjectPathTest(null, false);
        }

        [TestMethod]
        public void IsNotValidIfProjectPathPointsToFileThatDoesNotExistTest()
        {
            PerformProjectPathTest(s_tempFilePath + "1", false);
        }

        [TestMethod]
        public void IsValidIfProjectPathPointsToFileThatExistsTest()
        {
            PerformProjectPathTest(s_tempFilePath, true);
        }

        private void PerformProjectPathTest(string path, bool expectedIsValid)
        {
            var r = Create(1);
            r.UsesInput = false;
            r.UsesOutput = false;
            r.ProjectPath = path;
            Assert.AreEqual(expectedIsValid, r.IsValid);
        }
    }
}