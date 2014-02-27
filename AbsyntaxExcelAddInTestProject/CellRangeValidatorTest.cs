using System;
using AbsyntaxExcelAddIn.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AbsyntaxExcelAddInTestProject
{
    [TestClass]
    public class CellRangeValidatorTest
    {
        private static CellRangeValidator Create(string range)
        {
            var mock = new Mock<INamedRangeProvider>();
            var provider = mock.Object;
            return Create(range, provider);
        }

        private static CellRangeValidator Create(string range, INamedRangeProvider provider)
        {
            return new CellRangeValidator(range, provider);
        }

        [TestMethod]
        public void TopLeftWorksheetCellIsValidTest()
        {
            var v = Create("A1:A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void BottomRightWorksheetCellIsValidTest()
        {
            var v = Create("XFD1048576:XFD1048576");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void SingleCellNotationIsValidTest()
        {
            var v = Create("A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void LowerCaseIsValidTest()
        {
            var v = Create("a1:b1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void EntireSheetIsValidTest()
        {
            var v = Create("A1:XFD1048576");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void CellOrdersCanBeReversedTest()
        {
            var v = Create("XFD1048576:A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void LessThanMinimimRowIsInvalidTest()
        {
            var v = Create("A0:A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void GreaterThanMaximimRowIsInvalidTest()
        {
            var v = Create("A1048577:A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void GreaterThanMaximimColumnIsInvalidTest()
        {
            var v = Create("A1:XFE1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void SingleCellWithColonIsInvalidTest()
        {
            var v = Create("A1:");
            Assert.IsFalse(v.IsValid);
            v = Create(":A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void SpacesAreInvalidTest()
        {
            var v = Create("A 1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void MissingRowIsInvalidTest()
        {
            var v = Create("A");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void MissingColumnIsInvalidTest()
        {
            var v = Create("1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void GetRangeThrowsExceptionIfInvalidTest()
        {
            var v = Create("A0:A1");
            var ws = new Mock<Excel.Worksheet>().Object;
            v.GetRange(ws);
        }

        [TestMethod]
        public void RangeNameIsValidIfItAssociatesWithWorksheetTest()
        {
            var mock = new Mock<INamedRangeProvider>();
            var ws = new Mock<Excel.Worksheet>().Object;
            mock.Setup(m => m.IdentifyWorksheet(It.IsAny<string>())).Returns(ws);
            var v = Create(String.Empty, mock.Object);
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void RangeNameIsInvalidIfItDoesNotAssociateWithWorksheetTest()
        {
            var mock = new Mock<INamedRangeProvider>();
            mock.Setup(m => m.IdentifyWorksheet(It.IsAny<string>())).Returns(null as Excel.Worksheet);
            var v = Create(String.Empty, mock.Object);
            Assert.IsFalse(v.IsValid);
        }
    }
}