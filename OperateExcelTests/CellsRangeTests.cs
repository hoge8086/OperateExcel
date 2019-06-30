using Microsoft.VisualStudio.TestTools.UnitTesting;
using OperateExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OperateExcel.Tests
{
    [TestClass()]
    public class CellsRangeTests
    {

        [TestMethod()]
        public void CellByR1C1Test_A1()
        {
            var range = Cell.R1C1(1, 1);
            Assert.AreEqual("A1", range.A1Address);
        }

        [TestMethod()]
        public void CellByR1C1Test_Z1()
        {
            var range = Cell.R1C1(1, 26);
            Assert.AreEqual("Z1", range.A1Address);
        }
        [TestMethod()]
        public void CellByR1C1Test_AA1()
        {
            var range = Cell.R1C1(1, 27);
            Assert.AreEqual("AA1", range.A1Address);
        }
        [TestMethod()]
        public void CellByR1C1Test_BA1()
        {
            var range = Cell.R1C1(1, 53);
            Assert.AreEqual("BA1", range.A1Address);
        }

        [TestMethod()]
        public void CellByA1ColumnTest()
        {
            var range = Cell.A1(100, "AB");
            Assert.AreEqual("AB100", range.A1Address);
        }
    }
}