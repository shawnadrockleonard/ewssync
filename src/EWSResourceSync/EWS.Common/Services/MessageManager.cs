using EWS.Common.Models;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.ServiceBus.Messaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class MessageManager
    {
        public MessageManager(CancellationToken token)
        {
            _cancel = token;
            _sendFromO365Queue = QueueClient.CreateFromConnectionString(EWSConstants.Config.ServiceBus.SendFromO365);
        }


        CancellationToken _cancel;
        QueueClient _sendFromO365Queue;
        static string[] Actions = new string[] { "Created", "Deleted", "Modified" };

        public async System.Threading.Tasks.Task SendFromO365Async(string roomSMPT, Appointment apt, int status)
        {
            Trace.WriteLine($"SendFromO365Async({roomSMPT}, {apt.Subject}) starting");

            var refNoProp = apt.ExtendedProperties.FirstOrDefault(p => (p.PropertyDefinition == EWSConstants.RefIdPropertyDef));

            var meetingKeyProp = apt.ExtendedProperties.FirstOrDefault(p => (p.PropertyDefinition == EWSConstants.MeetingKeyPropertyDef));

            var booking = new UpdatedBooking()
            {
                EndUTC = apt.End.ToString(),
                StartUTC = apt.Start.ToString(),
                Subject = apt.Subject,
                BookingRef = refNoProp != null ? refNoProp.Value.ToString() : String.Empty,
                SiteMailBox = roomSMPT,
                Location = apt.Location,
                MeetingKey = meetingKeyProp != null ? (int)meetingKeyProp.Value : 0,
                CancelStatus = status
            };

            Trace.WriteLine($"{Actions[status]}: {booking.Subject}");
            await _sendFromO365Queue.SendAsync(new BrokeredMessage(booking));
            await System.Threading.Tasks.Task.FromResult(0);
            Trace.WriteLine($"SendFromO365Async() exiting");
        }

        public event Action<UpdatedBooking> NewBookingToO365;

        public async System.Threading.Tasks.Task StartGetToO365Async()
        {
            Trace.WriteLine($"GetToO365() starting");
            await System.Threading.Tasks.Task.Run(async () =>
            {
                var queueClient = QueueClient.CreateFromConnectionString(EWSConstants.Config.ServiceBus.ReadToO365);
                try
                {
                    Trace.WriteLine($"GetToO365.OnMessageAsync starting");
                    queueClient.OnMessage((msg) =>
                    {
                        Trace.WriteLine($"GetToO365 message received");
                        var booking = msg.GetBody<UpdatedBooking>();
                        if (NewBookingToO365 != null)
                        {
                            Trace.WriteLine($"Booking message {booking.Subject} insert into anon-async NewBooking.");
                            NewBookingToO365(booking);
                        }
                    }, new OnMessageOptions() { AutoComplete = true, MaxConcurrentCalls = 5 });
                    _cancel.ThrowIfCancellationRequested();
                    Console.ReadKey();
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Failed in Task=>StartGetToO365Async {ex.Message}");
                }
                await queueClient.CloseAsync();
            });
            Trace.WriteLine($"GetToO365() exiting");
        }
    }
}
