using EWS.Common;
using EWS.Common.Models;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;

namespace EWSResourceSync
{
    class Program
    {
        static AppSettings config;
        private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        MessageManager _queue;

        private static List<StreamingSubscriptionConnection> _connections = new List<StreamingSubscriptionConnection>();

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");

            config = AppSettings.Current;

            var p = new Program();
            p.Run();
        }

        static Program()
        {
        }

        private void Run()
        {
            _queue = new MessageManager(CancellationTokenSource.Token);
            _queue.NewBookingToO365 += async (b) => await ReserveRoomAsync(b);

            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                tasks.Add(_queue.StartGetToO365Async());
                tasks.Add(ListenToRoomReservationChangesAsync(EWService.Config.Exchange.ImpersonationAcct));

                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public class ItemEventComparer : IEqualityComparer<ItemEvent>
        {
            public bool Equals(ItemEvent x, ItemEvent y)
            {
                // If reference same object including null then return true
                if (object.ReferenceEquals(x, y))
                {
                    return true;
                }

                // If one object null the return false
                if (object.ReferenceEquals(x, null) || object.ReferenceEquals(y, null))
                {
                    return false;
                }

                // Compare properties for equality
                return (x.EventType == y.EventType && x.ItemId == y.ItemId);
            }

            public int GetHashCode(ItemEvent obj)
            {
                if (object.ReferenceEquals(obj, null))
                {
                    return 0;
                }

                int EventTypeHash = obj.EventType.GetHashCode();
                int ItemIdHash = obj.ItemId.GetHashCode();

                return EventTypeHash ^ ItemIdHash;
            }
        }

        // # TODO: Need to handle subscription for re-hydration
        // # TODO: evaluate why extension attributes are not returned in listener

        private async System.Threading.Tasks.Task ListenToRoomReservationChangesAsync(string mailboxOwner)
        {
            Trace.WriteLine($"ListenToRoomReservationChangesAsync({mailboxOwner}) starting");

            ServicePointManager.DefaultConnectionLimit = ServicePointManager.DefaultPersistentConnectionLimit;
            var service = await EWService.CreateExchangeServiceAsync();
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, EWService.Config.Exchange.ImpersonationAcct);

            var subs = new Dictionary<StreamingSubscription, ImpersonatedUserId>();
            var reqProps = EWService.extendedProperties;

            //TODO: Is there a better way to enumerate a la PS Get-MailBox | Where {$_.ResourceType -eq "Room"}
            foreach (var list in service.GetRoomLists())
            {
                foreach (var room in service.GetRooms(list))
                {
                    try
                    {
                        //TODO: Is there a more scalable way so we don't need to subscribe to each room individually?
                        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, room.Address);

                        //TODO: How to reconnect after app failure and get all events since failure occured
                        var sub = service.SubscribeToStreamingNotifications(
                            new FolderId[] { WellKnownFolderName.Calendar },
                            EventType.Created,
                            EventType.Deleted,
                            EventType.Modified,
                            EventType.Moved,
                            EventType.Copied);

                        Trace.WriteLine($"ListenToRoomReservationChangesAsync.Subscribed to room {room.Address}");
                        subs.Add(sub, service.ImpersonatedUserId);
                    }
                    catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException sex)
                    {
                        Trace.WriteLine($"Failed to provision subscription {sex.Message}");
                        throw new Exception($"Subscription could not be created for {room.Address} with MSG:{sex.Message}");
                    }
                }
            }

            // Create a streaming connection to the service object, over which events are returned to the client.
            // Keep the streaming connection open for 30 minutes.
            var connection = new StreamingSubscriptionConnection(service, subs.Keys.Select(s => s), 30);
            var semaphore = new System.Threading.SemaphoreSlim(1);

            connection.OnNotificationEvent += async (s, a) =>
            {
                Trace.WriteLine($"ListenToRoomReservationChangesAsync received notification");
                // New appt: ev.Type == Created
                // Del apt:  ev.Type == Moved, IsCancelled == true
                // Move apt: ev.Type == Modified, IsUnmodified == false, 
                var evGroups = a.Events.Where(ev => ev is ItemEvent)
                    .Select(ev => ((ItemEvent)ev))
                    .OrderBy(ev => ev.ItemId.UniqueId)
                    .GroupBy(ev => ev.ItemId);

                var filterPropertySet = new PropertySet(
                        AppointmentSchema.Location,
                        AppointmentSchema.Subject,
                        AppointmentSchema.Start,
                        AppointmentSchema.End,
                        AppointmentSchema.IsMeeting,
                        AppointmentSchema.IsOnlineMeeting,
                        AppointmentSchema.IsAllDayEvent,
                        AppointmentSchema.IsRecurring,
                        AppointmentSchema.IsCancelled,
                        AppointmentSchema.IsUnmodified,
                        AppointmentSchema.TimeZone,
                        AppointmentSchema.ICalUid,
                        AppointmentSchema.ParentFolderId,
                        AppointmentSchema.ConversationId,
                        AppointmentSchema.ICalRecurrenceId,
                        EWService.RefIdPropertyDef,
                        EWService.MeetingKeyPropertyDef);

                foreach (var evGroup in evGroups)
                {
                    Appointment appointmentTime = null;
                    var itemId = evGroup.Key;

                    var unfilteredEvents = evGroup.ToList();
                    var filterEvents = unfilteredEvents.Distinct(new ItemEventComparer()).OrderByDescending(x => x.TimeStamp);
                    foreach (ItemEvent ev in filterEvents)
                    {
                        // Find an item event you can bind to
                        int action = 99;
                        var ewsService = a.Subscription.Service;
                        var subscription = subs.FirstOrDefault(k => k.Key.Id == a.Subscription.Id);

                        try
                        {
                            ewsService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, subscription.Value.Id);
                            appointmentTime = (Appointment)Item.Bind(ewsService, itemId);



                            if (ev.EventType == EventType.Created)
                            {
                                action = 0; // created
                            }
                            else if (ev.EventType == EventType.Moved && appointmentTime.IsCancelled)
                            {
                                action = 1; // deleted
                            }
                            else if (ev.EventType == EventType.Modified && !appointmentTime.IsUnmodified)
                            {
                                action = 2; // modified
                            }
                            else
                            {
                                continue;
                            }

                            ExtendedPropertyDefinition CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
                            PropertySet psPropSet = new PropertySet(BasePropertySet.FirstClassProperties)
                            {
                                CleanGlobalObjectId
                            };
                            appointmentTime.Load(psPropSet);
                            appointmentTime.TryGetProperty(CleanGlobalObjectId, out object CalIdVal);

                            var icalId = appointmentTime.ICalUid;
                            var mailboxId = appointmentTime.Organizer.Address;
                            ewsService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxId);
                            CalendarFolder AtndCalendar = CalendarFolder.Bind(ewsService, new FolderId(WellKnownFolderName.Calendar, mailboxId), filterPropertySet);
                            SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(CleanGlobalObjectId, Convert.ToBase64String((Byte[])CalIdVal));
                            ItemView ivItemView = new ItemView(5)
                            {
                                PropertySet = filterPropertySet
                            };
                            FindItemsResults<Item> fiResults = AtndCalendar.FindItems(sfSearchFilter, ivItemView);
                            if (fiResults.Items.Count > 0)
                            {
                                var filterApt = fiResults.Items.FirstOrDefault() as Appointment;
                                Trace.WriteLine($"The first {fiResults.Items.Count()} appointments on your calendar from {filterApt.Start.ToShortDateString()} to {filterApt.End.ToShortDateString()}");

                                var props = filterApt.ExtendedProperties.Where(p => (p.PropertyDefinition.PropertySet == DefaultExtendedPropertySet.Meeting));
                                string refId = string.Empty;
                                int meetingKey = 0;
                                if (props.Count() > 0)
                                {
                                    refId = (string)props.First(p => p.PropertyDefinition.Name == EWService.RefIdPropertyName).Value;
                                    meetingKey = (int)props.First(p => p.PropertyDefinition.Name == EWService.MeetingKeyPropertyName).Value;
                                }
                            }
                        }
                        catch (ServiceResponseException ex)
                        {
                            Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                            continue;
                        }

                        try
                        {

                            var mailboxId = appointmentTime.Organizer.Address;
                            // Initialize the calendar folder via Impersonation
                            ewsService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxId);
                            // Initialize the calendar folder object with only the folder ID. 
                            var cfFolderId = new FolderId(WellKnownFolderName.Calendar, mailboxId);

                            // Set the start and end time and number of appointments to retrieve.
                            CalendarFolder calendar = CalendarFolder.Bind(ewsService, cfFolderId, new PropertySet());

                            // Limit the properties returned to the appointment's subject, start time, and end time.
                            CalendarView cView = new CalendarView(appointmentTime.Start, appointmentTime.End, EWService.Config.Exchange.BatchSize)
                            {
                                PropertySet = filterPropertySet
                            };

                            // Retrieve a collection of appointments by using the calendar view.
                            FindItemsResults<Appointment> calAppointments = calendar.FindAppointments(cView);
                            Trace.WriteLine($"The first {calAppointments.Count()} appointments on your calendar from {appointmentTime.Start.ToShortDateString()} to {appointmentTime.End.ToShortDateString()}");

                            if (!calAppointments.Any())
                            {
                                Trace.WriteLine("This apt has no extended properties!!!");
                            }
                            else
                            {
                                var apt = calAppointments.FirstOrDefault();
                                var props = apt.ExtendedProperties.Where(p => (p.PropertyDefinition.PropertySet == DefaultExtendedPropertySet.Meeting));
                                string refId = string.Empty;
                                int meetingKey = 0;
                                if (props.Count() > 0)
                                {
                                    refId = (string)props.First(p => p.PropertyDefinition.Name == EWService.RefIdPropertyName).Value;
                                    meetingKey = (int)props.First(p => p.PropertyDefinition.Name == EWService.MeetingKeyPropertyName).Value;
                                }
                            }
                        }
                        catch (ServiceResponseException ex)
                        {
                            Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                            continue;
                        }

                        await _queue.SendFromO365Async(ewsService.ImpersonatedUserId.Id, appointmentTime, action);
                    }
                }
            };
            connection.OnDisconnect += (s, a) =>
            {
                if (a.Exception == null)
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection disconnected");
                else
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection disconnected with exception: {a.Exception.Message}");
                if (CancellationTokenSource.IsCancellationRequested)
                    semaphore.Release();
                else
                {
                    connection.Open();
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection re-connected");
                }
            };
            connection.OnSubscriptionError += (sender, args) =>
            {
                if (args.Exception is ServiceResponseException)
                {
                    var exception = args.Exception as ServiceResponseException;
                }
                else if (args.Exception != null)
                {
                    Trace.WriteLine($"OnSubscriptionError() : {args.Exception.Message} Stack Trace : {args.Exception.StackTrace} Inner Exception : {args.Exception.InnerException}");
                }

                connection = (StreamingSubscriptionConnection)sender;
                if (args.Subscription != null)
                {
                    try
                    {
                        connection.RemoveSubscription(args.Subscription);
                    }
                    catch (Exception rex)
                    {
                        Trace.WriteLine($"RemoveSubscriptionException to provision subscription {rex.Message}");
                    }
                }
            };
            semaphore.Wait();
            connection.Open();
            await semaphore.WaitAsync(CancellationTokenSource.Token);

            Trace.WriteLine($"ListenToRoomReservationChangesAsync({mailboxOwner}) exiting");
        }

        private void OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            string fromEmailAddress = args.Subscription.Service.ImpersonatedUserId.Id;
        }

        private async System.Threading.Tasks.Task ReserveRoomAsync(UpdatedBooking booking)
        {
            Trace.WriteLine($"ReserveRoomAsync({booking.Location}) starting");
            var service = await EWService.CreateExchangeServiceAsync();
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, booking.MailBoxOwnerEmail);
            Appointment meeting = new Appointment(service);
            meeting.Resources.Add(booking.SiteMailBox);
            meeting.Subject = booking.Subject;
            //appointment.Body = "...";
            meeting.Start = DateTime.Parse(booking.StartUTC);
            meeting.End = DateTime.Parse(booking.EndUTC);
            meeting.Location = booking.Location;
            meeting.SetExtendedProperty(EWService.RefIdPropertyDef, booking.BookingRef);
            meeting.SetExtendedProperty(EWService.MeetingKeyPropertyDef, booking.MeetingKey);
            //meeting.ReminderDueBy = DateTime.Now;
            meeting.Save(SendInvitationsMode.SendOnlyToAll);
            // Verify that the appointment was created by using the appointment's item ID.
            var item = Item.Bind(service, meeting.Id, new PropertySet(ItemSchema.Subject, EWService.RefIdPropertyDef, EWService.MeetingKeyPropertyDef));
            Console.WriteLine($"Appointment created: {item.Subject}");
            Trace.WriteLine($"ReserveRoomAsync({booking.Location}) completed");
            //TODO: The room will may decline if booked. Always declines for past dates. See also: https://social.msdn.microsoft.com/Forums/exchange/en-US/cead7451-dcc5-46b9-b225-b16874fdc914/ews-confirming-room-response-accepteddeclined-when-creating-appointment-where-room-is-invited
        }




        private async System.Threading.Tasks.Task GetRooms()
        {
            Trace.WriteLine("GetRooms starting");
            var service = await EWService.CreateExchangeServiceAsync();

            try
            {
                foreach (var list in service.GetRoomLists())
                {
                    Console.WriteLine("List: {0} - {1}", list.Name, list.Address);
                    foreach (var room in service.GetRooms(list))
                    {
                        Console.WriteLine("   Room: {0} - {1}", room.Name, room.Address);
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("GetRooms exception: {0}", ex.Message);
                throw;
            }
            Trace.WriteLine("GetRooms completed");
        }
    }
}
