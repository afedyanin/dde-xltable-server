namespace DdeExcelTableServer.Data
{
    using System;
    using System.Text;

    public static class DdeExtensions
    {
        public static string RowToString(this DdeTable table, int row)
        {
            return table.RowToString(table.GetRow(row));
        }

        public static string RowToString(this DdeTable table, object[] row)
        {
            var sb = new StringBuilder();

            foreach (var k in row)
            {
                sb.Append($"{k} ");
            }

            return sb.ToString();
        }

        public static string AllRowsToString(this DdeTable table, int maxRows = 0)
        {
            var sb = new StringBuilder();

            var rows = maxRows == 0 ? table.Rows : Math.Min(maxRows, table.Rows);

            for (var i = 0; i < rows; i++)
            {
                var row = table.GetRow(i);

                foreach (var k in row)
                {
                    sb.Append($"{k} ");
                }

                if (i < rows - 1)
                {
                    sb.Append("\n");
                }
            }

            return sb.ToString();
        }
    }
}
