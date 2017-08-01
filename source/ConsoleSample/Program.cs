namespace ConsoleSample
{
    using System;
    using DdeExcelTableServer;
    using DdeExcelTableServer.Data;

    class Program
    {
        static void Main(string[] args)
        {
            using (var ddeServer = new Server("Vert"))
            {
                ddeServer.OnPokeAction = (topic, item, data, format) =>
                {
                    var message = new DdeMessage(data, topic, item); // Excel table format by default
                    Console.WriteLine($"{message.Topic}|{message.Item}: {message.Table.AllRowsToString()}");
                };

                Console.WriteLine("Starting server... Press <Enter> to exit");
                ddeServer.Register();

                Console.ReadLine();
                ddeServer.Unregister();
            }
        }
    }
}
