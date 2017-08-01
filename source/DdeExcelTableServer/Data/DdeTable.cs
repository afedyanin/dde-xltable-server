namespace DdeExcelTableServer.Data
{
    using System;
    using System.Collections.Generic;

    public class DdeTable
    {
        public int Rows { get; }
        public int Cols { get; }
        public object[] Cells { get; }

        public DdeTable(int rows, int cols, object[] cells)
        {
            this.Rows = rows;
            this.Cols = cols;
            this.Cells = cells;
        }

        public object GetCell(int row, int col)
        {
            this.ThrowIfRowInvalid(row);
            this.ThrowIfColInvalid(col);
            return this.GetCellInternal(row, col);
        }
        public object[] GetRow(int row)
        {
            this.ThrowIfRowInvalid(row);
            return this.GetRowInternal(row);
        }
        public IList<object[]> GetRows()
        {
            var res = new List<object[]>(this.Rows);

            for (var row = 0; row < this.Rows; row++)
            {
                res.Add(this.GetRowInternal(row));
            }

            return res;
        }

        private object GetCellInternal(int row, int col)
        {
            return this.Cells[row * this.Cols + col];
        }
        private object[] GetRowInternal(int row)
        {
            var res = new object[this.Cols];

            for (var col = 0; col < this.Cols; col++)
            {
                res[col] = this.GetCellInternal(row, col);
            }

            return res;
        }
        private void ThrowIfRowInvalid(int row)
        {
            if (row < 0)
            {
                throw new InvalidOperationException("Row must be >= 0");
            }

            var maxRow = this.Rows - 1;

            if (row > maxRow)
            {
                throw new InvalidOperationException($"Row must be <= {maxRow}");
            }
        }
        private void ThrowIfColInvalid(int col)
        {
            if (col < 0)
            {
                throw new InvalidOperationException("Column must be >= 0");
            }

            var maxCol = this.Cols - 1;

            if (col > maxCol)
            {
                throw new InvalidOperationException($"Column must be <= {maxCol}");
            }
        }
    }
}
