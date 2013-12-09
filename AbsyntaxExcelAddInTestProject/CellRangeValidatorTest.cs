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
        [TestMethod]
        public void TopLeftWorksheetCellIsValidTest()
        {
            var v = new CellRangeValidator("A1:A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void BottomRightWorksheetCellIsValidTest()
        {
            var v = new CellRangeValidator("XFD1048576:XFD1048576");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void SingleCellNotationIsValidTest()
        {
            var v = new CellRangeValidator("A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void LowerCaseIsValidTest()
        {
            var v = new CellRangeValidator("a1:b1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void EntireSheetIsValidTest()
        {
            var v = new CellRangeValidator("A1:XFD1048576");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void CellOrdersCanBeReversedTest()
        {
            var v = new CellRangeValidator("XFD1048576:A1");
            Assert.IsTrue(v.IsValid);
        }

        [TestMethod]
        public void LessThanMinimimRowIsInvalidTest()
        {
            var v = new CellRangeValidator("A0:A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void GreaterThanMaximimRowIsInvalidTest()
        {
            var v = new CellRangeValidator("A1048577:A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void GreaterThanMaximimColumnIsInvalidTest()
        {
            var v = new CellRangeValidator("A1:XFE1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void SingleCellWithColonIsInvalidTest()
        {
            var v = new CellRangeValidator("A1:");
            Assert.IsFalse(v.IsValid);
            v = new CellRangeValidator(":A1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void SpacesAreInvalidTest()
        {
            var v = new CellRangeValidator("A 1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void MissingRowIsInvalidTest()
        {
            var v = new CellRangeValidator("A");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void MissingColumnIsInvalidTest()
        {
            var v = new CellRangeValidator("1");
            Assert.IsFalse(v.IsValid);
        }

        [TestMethod]
        public void FirstAndLastCellsAreNullWhenInvalidTest()
        {
            var v = new CellRangeValidator("A0:A1");
            Assert.IsNull(v.FirstCell);
            Assert.IsNull(v.LastCell);
        }

        [TestMethod]
        public void FirstAndLastCellsAreAsSpecifiedWhenValidTest()
        {
            string c1 = "A1";
            string c2 = "B2";
            var v = new CellRangeValidator(String.Format("{0}:{1}", c1, c2));
            Assert.AreEqual(c1, v.FirstCell);
            Assert.AreEqual(c2, v.LastCell);
        }

        [TestMethod]
        public void FirstAndLastCellsAreEqualForValidSingleCellNotationTest()
        {
            string c = "A1";
            var v = new CellRangeValidator(String.Format("{0}:{0}", c));
            Assert.AreEqual(c, v.FirstCell);
            Assert.AreEqual(c, v.LastCell);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void GetRangeThrowsExceptionIfInvalidTest()
        {
            var v = new CellRangeValidator("A0:A1");
            var ws = new Mock<Excel.Worksheet>().Object;
            v.GetRange(ws);
        }
    }
}