using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OperateExcel
{
    public class CellsRange
    {
        public string A1Address { get; private set; }

        protected CellsRange(string a1Address)
        {
            this.A1Address = a1Address;
        }

        public static CellsRange A1(string a1)
        {
            return new CellsRange(a1);
        }

        public static CellsRange ColumnRange(string column)
        {
            return ColumnRange(column, column);
        }
        public static CellsRange ColumnRange(string start, string end)
        {
            return new CellsRange(start + ":" + end);
        }
    }

    public class Cell : CellsRange
    {
        private Cell(string a1Address) : base(a1Address) { }

        public static Cell R1C1(int row, int column)
        {
            if (row < 1 || column < 1)
                throw new ArgumentException("列番号あるは行番号が範囲外です.");
            return A1(row, ToA1Column(column));
        }

        public static new Cell A1(string a1)
        {
            if (!Regex.IsMatch(a1, @"^[a-zA-Z]+\d+$"))
                throw new ArgumentException(a1 +"は一つのセルのみを参照する必要があります.");

            return new Cell(a1);
        }

        public static Cell A1(int row, string a1Column)
        {
            if (row < 1)
                throw new ArgumentException("行番号が範囲外です.");
            return new Cell(a1Column + row);
        }

        private static string ToA1Column(int columnIndex)
        {
            if (columnIndex < 1)
                throw new ArgumentException("列番号が範囲外です.");
            return ToA1ColumnImpl(columnIndex);
        }

        private static string ToA1ColumnImpl(int columnIndex)
        {
            if (columnIndex < 1) return string.Empty;
            return ToA1ColumnImpl((columnIndex- 1) / 26) + (char)('A' + ((columnIndex- 1) % 26));
        }
    }

}
