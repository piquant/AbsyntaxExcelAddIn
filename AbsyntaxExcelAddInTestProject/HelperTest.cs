using System;
using AbsyntaxExcelAddIn.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject
{
    [TestClass]
    public class HelperTest
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void GetCellThrowsExceptionIfColIndexLessThanOneTest()
        {
            Helper.GetCell(0, 1);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void GetCellThrowsExceptionIfRowLessThanOneTest()
        {
            Helper.GetCell(1, 0);
        }

        [TestMethod]
        public void GetCellReturnsExpectedValueTest()
        {
            var items = new Tuple<int, int, string>[] {
                new Tuple<int, int, string>(1, 1, "A1"),
                new Tuple<int, int, string>(2, 1, "B1"),
                new Tuple<int, int, string>(1, 2, "A2"),
                new Tuple<int, int, string>(27, 1, "AA1"),
                new Tuple<int, int, string>(28, 1, "AB1"),
                new Tuple<int, int, string>(53, 1, "BA1"),
                new Tuple<int, int, string>(16384, 1048576, "XFD1048576")
            };
            PerformGetCellTest(items);
        }

        private void PerformGetCellTest(Tuple<int, int, string>[] items)
        {
            foreach (Tuple<int, int, string> item in items) {
                string cell = Helper.GetCell(item.Item1, item.Item2);
                Assert.AreEqual(item.Item3, cell);
            }
        }

        [TestMethod]
        public void CreateIdReturnsNumberNotInListTest()
        {
            int[] ids = { 1, 2, 3, 5 };
            int id = Helper.CreateId(ids);
            Assert.AreEqual(4, id);
        }

        [TestMethod]
        public void GetMillisecondsReturnsExpectedValueTest()
        {
            int oneSec = 1000;
            var items = new Tuple<int, TimeUnit, int>[] {
                new Tuple<int, TimeUnit, int>(1, TimeUnit.Seconds, oneSec),
                new Tuple<int, TimeUnit, int>(2, TimeUnit.Seconds, oneSec * 2),
                new Tuple<int, TimeUnit, int>(1, TimeUnit.Minutes, oneSec * 60),
                new Tuple<int, TimeUnit, int>(2, TimeUnit.Minutes, oneSec * 60 * 2),
                new Tuple<int, TimeUnit, int>(1, TimeUnit.Hours, oneSec * 60 * 60),
                new Tuple<int, TimeUnit, int>(2, TimeUnit.Hours, oneSec * 60 * 60 * 2),
                new Tuple<int, TimeUnit, int>(1, TimeUnit.Days, oneSec * 60 * 60 * 24),
                new Tuple<int, TimeUnit, int>(2, TimeUnit.Days, oneSec * 60 * 60 * 24 * 2)
            };
            PerformGetMillisecondsTest(items);
        }

        private void PerformGetMillisecondsTest(Tuple<int, TimeUnit, int>[] items)
        {
            foreach (Tuple<int, TimeUnit, int> item in items) {
                int ms = Helper.GetMilliseconds(item.Item1, item.Item2);
                Assert.AreEqual(item.Item3, ms);
            }
        }
    }
}