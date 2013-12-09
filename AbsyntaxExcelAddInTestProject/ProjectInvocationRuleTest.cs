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

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullWorksheetProviderThrowsException1Test()
        {
            new ProjectInvocationRule(null, 1);
        }
        
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void NullWorksheetProviderThrowsException2Test()
        {
            var dr = new Mock<IDataReader>().Object;
            new ProjectInvocationRule(null, dr);
        }

        [TestMethod]
        public void DeserialisationConstructorTest()
        {
            var wp = new Mock<IWorksheetProvider>().Object;
            var drMock = new Mock<IDataReader>();
            var r = new ProjectInvocationRule(wp, drMock.Object);
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
            var wp = new Mock<IWorksheetProvider>().Object;
            var r = new ProjectInvocationRule(wp, 1);
            Assert.IsNull(r.InputSheetKey);
        }

        [TestMethod]
        public void InputSheetNameIsNullIfNoSheetsToSelectTest()
        {
            var wp = new Mock<IWorksheetProvider>().Object;
            var r = new ProjectInvocationRule(wp, 1);
            Assert.IsNull(r.InputSheetName);
        }

        [TestMethod]
        public void OutputSheetKeyIsNullIfNoSheetsToSelectTest()
        {
            var wp = new Mock<IWorksheetProvider>().Object;
            var r = new ProjectInvocationRule(wp, 1);
            Assert.IsNull(r.OutputSheetKey);
        }

        [TestMethod]
        public void OutputSheetNameIsNullIfNoSheetsToSelectTest()
        {
            var wp = new Mock<IWorksheetProvider>().Object;
            var r = new ProjectInvocationRule(wp, 1);
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
            var wp = new Mock<IWorksheetProvider>().Object;
            var r = new ProjectInvocationRule(wp, 1);
            r.UsesInput = false;
            r.UsesOutput = false;
            r.ProjectPath = path;
            Assert.AreEqual(expectedIsValid, r.IsValid);
        }
    }
}