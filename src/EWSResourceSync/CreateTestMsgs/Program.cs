using EWS.Common;
using EWS.Common.Services;
using Microsoft.ServiceBus.Messaging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EWSResourcePush
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine();
            Console.Write("Send to O365? (y/n)?");
            var key = Console.ReadKey();
            if ((key.KeyChar == 'y') || (key.KeyChar == 'Y'))
            {
                Console.WriteLine();
                Console.Write("Sending...");
                var connStr = EWSConstants.Config.ServiceBus.SendToO365;
                var queue = QueueClient.CreateFromConnectionString(connStr);

                var str = File.OpenText("bookings.json");
                var inp = str.ReadToEnd();
                var bookings = JsonConvert.DeserializeObject<IEnumerable<EWS.Common.Models.UpdatedBooking>>(inp);

                var i = 0;
                foreach (var booking in bookings)
                {
                    var msg = new BrokeredMessage(booking);
                    queue.Send(msg);
                    Console.WriteLine("{0}. Sent: {1}", ++i, booking.Subject);
                }
                Console.WriteLine();
                Console.WriteLine("Sent {0} messages", i);
            }

            Console.WriteLine();
            Console.Write("Remove msgs to O365? (y/n)?");
            key = Console.ReadKey();
            if ((key.KeyChar == 'y') || (key.KeyChar == 'Y'))
            {
                Console.WriteLine();
                Console.WriteLine("Listening...");
                var queueClient = QueueClient.CreateFromConnectionString(EWSConstants.Config.ServiceBus.ReadToO365);
                queueClient.OnMessage((msg) =>
                {
                    var booking = msg.GetBody<EWS.Common.Models.UpdatedBooking>();
                    Console.WriteLine($"MailBoxOwnerEmail {booking.MailBoxOwnerEmail}");
                    msg.Complete();

                }, new OnMessageOptions() { AutoComplete = true, MaxConcurrentCalls = 5 });
                queueClient.Close();
            }

            Console.WriteLine();
            Console.Write("Get msgs from O365? (y/n)?");
            key = Console.ReadKey();
            if ((key.KeyChar == 'y') || (key.KeyChar == 'Y'))
            {
                Console.WriteLine();
                Console.WriteLine("Listening...");
                var queueClient = QueueClient.CreateFromConnectionString(EWSConstants.Config.ServiceBus.ReadFromO365);
                queueClient.OnMessage((msg) =>
                {
                    var booking = msg.GetBody<EWS.Common.Models.UpdatedBooking>();
                    Console.WriteLine($"Msg received: {booking.SiteMailBox} - {booking.Subject}. Cancel status: {booking.CancelStatus}");
                }, new OnMessageOptions() { AutoComplete = true, MaxConcurrentCalls = 5 });
                queueClient.Close();
            }

            Console.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }
    }
}
