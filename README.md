# dde-xltable-server

Simple .NET Dynamic Data Exchange (DDE) Server 

Based on NDde project: https://ndde.codeplex.com/
This library provides a convenient and easy way to integrate .NET applications with legacy applications that use Dynamic Data Exchange (DDE). DDE is an older interprocess communication protocol that relies heavily on passing windows messages back and forth between applications. Other, more modern and robust, techniques for interprocess communication are available and should be used when available. This library is only intended to be used when no other alternatives exist.

Sample:

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


[NuGet package](https://www.nuget.org/packages/DdeExcelTableServer)

To install DdeExcelTableServer, run the following command in the Package Manager Console


    PM> Install-Package DdeExcelTableServer
