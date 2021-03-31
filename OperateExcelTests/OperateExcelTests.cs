using Microsoft.VisualStudio.TestTools.UnitTesting;
using OperateExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OperateExcel.Tests
{
    [TestClass()]
    public class OperateExcelTests
    {
        private string TestDir = "test";
        private string CalcPath(string filePath)
        {
            return TestDir + @"\" + filePath;
        }

        [TestInitialize]
        public void TestInitialize()
        {
            if (!Directory.Exists(TestDir))
                Directory.CreateDirectory(TestDir);

            //ディレクトリ以外の全ファイルを削除
            string[] filePaths = Directory.GetFiles(TestDir);
            foreach (string filePath in filePaths)
            {
                File.SetAttributes(filePath, FileAttributes.Normal);
                File.Delete(filePath);
            }
        }

        [TestMethod()]
        public void OpenTest_Create()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open();
                excel.Write(Cell.R1C1(1, 1), "1", "2", "3");
                excel.Write(Cell.R1C1(101, 101), "あいうえおかきくけこ");
                excel.Save(CalcPath("test.xlsx"));
            }

            Assert.IsTrue(File.Exists(CalcPath("test.xlsx")));

            using (var excel = new OperateExcel())
            {
                excel.Open(CalcPath("test.xlsx"));
                Assert.AreEqual(1, excel.Read<int>(Cell.R1C1(1, 1)));
                Assert.AreEqual("2", excel.Read<string>(Cell.R1C1(1, 2)));
                Assert.AreEqual(3.0, excel.Read<double>(Cell.R1C1(1, 3)));
                Assert.AreEqual("あいうえおかきくけこ", excel.Read<string>(Cell.R1C1(101, 101)));

                excel.Write(Cell.R1C1(4, 4), "さしすせそ");
                excel.Save();
            }

            using (var excel = new OperateExcel())
            {
                excel.Open(CalcPath("test.xlsx"));
                Assert.AreEqual("さしすせそ", excel.Read<string>(Cell.R1C1(4, 4)));
            }
        }

        [TestMethod()]
        public void OpenTest_Open()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                excel.SelectSheet("テスト");
                Assert.AreEqual("ほげほげ", excel.Read<string>(Cell.A1("B2")));
            }
        }

        [TestMethod()]
        public void CopyAndInsertRowTest()
        {
            File.Copy("template.xlsx", CalcPath("test.xlsx"));
            using (var excel = new OperateExcel())
            {
                excel.Open(CalcPath("test.xlsx"));
                excel.CopyAndInsertRow(1, 2, 2, "Sheet2");
                excel.CopyAndInsertRow(2, 2, 1);
                Assert.AreEqual("1", excel.Read<string>(Cell.R1C1(2, 2)));
                Assert.AreEqual("4", excel.Read<string>(Cell.R1C1(2, 5)));
                Assert.AreEqual("あ", excel.Read<string>(Cell.R1C1(3, 2)));
                Assert.AreEqual("あ", excel.Read<string>(Cell.R1C1(4, 2)));
                Assert.AreEqual("え", excel.Read<string>(Cell.R1C1(3, 5)));
                Assert.AreEqual("え", excel.Read<string>(Cell.R1C1(4, 5)));
            }
        }

        [TestMethod()]
        public void FindTest_Table()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                int row, column;
                excel.SelectSheet("Sheet2");
                Assert.IsTrue(excel.Find(CellsRange.A1("B2:E5"), "え", out row, out column));
                Assert.AreEqual(4, row);
                Assert.AreEqual(5, column);
            }
        }

        [TestMethod()]
        public void FindTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                int row, column;
                excel.SelectSheet("検索");
                Assert.IsTrue(excel.Find(CellsRange.ColumnRange("C"), "中川義男", out row, out column));
                Assert.AreEqual(9, row);
                Assert.AreEqual(3, column);
                Assert.IsFalse(excel.Find(CellsRange.ColumnRange("C"), "山本", out row, out column));
                Assert.IsTrue(excel.Find(CellsRange.ColumnRange("C"), "山本", out row, out column, true));
                Assert.AreEqual(12, row);
                Assert.AreEqual(3, column);
            }
        }

        [TestMethod()]
        public void GetSheetNamesTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                var names = excel.GetSheetNames();
                Assert.AreEqual(4, names.Count);
                Assert.AreEqual("Sheet1", names[0]);
                Assert.AreEqual("テスト", names[1]);
                Assert.AreEqual("Sheet2", names[2]);
                Assert.AreEqual("検索", names[3]);
            }
        }

        [TestMethod()]
        public void CreateSheetTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open();
                excel.CreateSheet("first", OperateExcel.SheetPosition.First);
                var names = excel.GetSheetNames();
                Assert.AreEqual(2, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);

                excel.CreateSheet("last", OperateExcel.SheetPosition.Last);
                names = excel.GetSheetNames();
                Assert.AreEqual(3, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);
                Assert.AreEqual("last", names[2]);

                excel.CreateSheet("after", OperateExcel.SheetPosition.AfterCurrentSheet);
                names = excel.GetSheetNames();
                Assert.AreEqual(4, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);
                Assert.AreEqual("after", names[2]);
                Assert.AreEqual("last", names[3]);

                excel.CreateSheet("before", OperateExcel.SheetPosition.BeforeCurrentSheet);
                names = excel.GetSheetNames();
                Assert.AreEqual(5, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("before", names[1]);
                Assert.AreEqual("Sheet1", names[2]);
                Assert.AreEqual("after", names[3]);
                Assert.AreEqual("last", names[4]);
            }
        }

        [TestMethod()]
        public void CopySheetTestReal()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                excel.CopySheet("コピーした奴", "検索", OperateExcel.SheetPosition.First);
                excel.CreateSheet("新規作成奴", OperateExcel.SheetPosition.AfterCurrentSheet);
                excel.SelectSheet("新規作成奴");
                excel.Write(Cell.A1("A1"), "おおおお");
                excel.Save("testCopy.xlsx");
            }
        }

        [TestMethod()]
        public void CopySheetTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open();
                excel.Write(Cell.A1("B2"), "あいうえお");
                excel.CopySheet("first", "Sheet1", OperateExcel.SheetPosition.First);
                var names = excel.GetSheetNames();
                Assert.AreEqual(2, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);

                excel.CopySheet("last", "Sheet1", OperateExcel.SheetPosition.Last);
                names = excel.GetSheetNames();
                Assert.AreEqual(3, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);
                Assert.AreEqual("last", names[2]);

                excel.CopySheet("after", "Sheet1", OperateExcel.SheetPosition.AfterCurrentSheet);
                names = excel.GetSheetNames();
                Assert.AreEqual(4, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("Sheet1", names[1]);
                Assert.AreEqual("after", names[2]);
                Assert.AreEqual("last", names[3]);

                excel.CopySheet("before", "Sheet1", OperateExcel.SheetPosition.BeforeCurrentSheet);
                names = excel.GetSheetNames();
                Assert.AreEqual(5, names.Count);
                Assert.AreEqual("first", names[0]);
                Assert.AreEqual("before", names[1]);
                Assert.AreEqual("Sheet1", names[2]);
                Assert.AreEqual("after", names[3]);
                Assert.AreEqual("last", names[4]);

                foreach (var name in names)
                {
                    excel.SelectSheet(name);
                    Assert.AreEqual("あいうえお", excel.Read<string>(Cell.A1("B2")));
                }
            }
        }

        [TestMethod()]
        public void IsStrikethroughTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                //全て取り消し線なし
                Assert.AreEqual(false, excel.IsStrikethrough(Cell.A1("B9")));
                //セルごと取り消し線
                Assert.AreEqual(true, excel.IsStrikethrough(Cell.A1("B10")));
                //セル内の一部を取り消し線
                Assert.AreEqual(false, excel.IsStrikethrough(Cell.A1("B11")));
                //セル内のテキストを全て取り消し線
                Assert.AreEqual(true, excel.IsStrikethrough(Cell.A1("B12")));
                //改行以外を全て取り消し線
                Assert.AreEqual(false, excel.IsStrikethrough(Cell.A1("B13")));
                Assert.AreEqual(true, excel.IsStrikethrough(CellsRange.A1("B14:C14")));
                Assert.AreEqual(false, excel.IsStrikethrough(CellsRange.A1("B15:C15")));
            }
        }

        [TestMethod()]
        public void ReadRowTest()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                excel.SelectSheet("Sheet2");
                var row = excel.ReadRow(4, ColumnsRange.A1("B:F"));
                Assert.AreEqual("あ", row[0]);
                Assert.AreEqual("い", row[1]);
                Assert.AreEqual("う", row[2]);
                Assert.AreEqual("え", row[3]);
                Assert.AreEqual("", row[4]);
                Assert.AreEqual(5, row.Count);
            }
        }
        [TestMethod()]
        public void ReadRowTestOne()
        {
            using (var excel = new OperateExcel())
            {
                excel.Open("template.xlsx");
                excel.SelectSheet("Sheet2");
                var row = excel.ReadRow(4, ColumnsRange.A1("B:B"));
                Assert.AreEqual("あ", row[0]);
                Assert.AreEqual(1, row.Count);
            }
        }
    }
}