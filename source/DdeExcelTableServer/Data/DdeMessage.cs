namespace DdeExcelTableServer.Data
{
    public class DdeMessage
    {
        public string Topic { get; }
        public string Item { get; }
        public DdeTable Table { get; }

        public DdeMessage(byte[] data, string topic = "", string item = "")
        {
            this.Table = XlTableFormat.Read(data, (rows, cols, cells) => new DdeTable(rows, cols, cells));
            this.Topic = topic;
            this.Item = item;
        }

        public DdeMessage(DdeTable table, string topic = "", string item = "")
        {
            this.Table = table;
            this.Topic = topic;
            this.Item = item;
        }
    }
}
