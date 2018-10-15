using EWS.Common;
using EWS.Common.Database;
using EWS.Common.Models;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace EWSResourcePull
{
    class Program
    {
        const int timeout = 30;
        static bool exitSystem = false;
        static string mailboxOwner = "";
        MessageManager _queue;
        private static System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        private static List<SubscriptionCollection> subscriptions = new List<SubscriptionCollection>();
        private static EWSDbContext EwsDatabase { get; set; }
        private static bool IsDisposed { get; set; }
        private static EWService EwsService { get; set; }
        private static AuthenticationResult Ewstoken { get; set; }

        public static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");

            _handler += new EventHandler(ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);


            mailboxOwner = EWSConstants.Config.Exchange.ImpersonationAcct;

            EwsDatabase = new EWSDbContext(EWSConstants.Config.Database);

            var service = System.Threading.Tasks.Task.Run(async () =>
            {
                var resservice = await EWSConstants.AcquireTokenAsync();
                return resservice;
            }, CancellationTokenSource.Token);
            service.Wait();


            Ewstoken = service.Result;
            EwsService = new EWService(Ewstoken);


            var p = new Program();
            p.Run();

            //hold the console so it doesn’t run off the end
            Console.ReadKey();
        }


        static Program()
        {
        }


        private void Run()
        {
            _queue = new MessageManager(CancellationTokenSource.Token);


            try
            {
                var roomHasChanged = false;
                var roomlisting = EwsService.GetRoomListing();
                foreach (var roomlist in roomlisting)
                {
                    foreach (var room in roomlist.Value)
                    {
                        EntityRoomListRoom databaseRoom = null;
                        if (EwsDatabase.RoomListRoomEntities.Any(s => s.SmtpAddress == room.Address))
                        {
                            databaseRoom = EwsDatabase.RoomListRoomEntities.FirstOrDefault(f => f.SmtpAddress == room.Address);
                        }
                        else
                        {
                            databaseRoom = new EntityRoomListRoom()
                            {
                                SmtpAddress = room.Address,
                                Identity = room.Address,
                                RoomList = roomlist.Key
                            };
                            EwsDatabase.RoomListRoomEntities.Add(databaseRoom);
                            roomHasChanged = true;
                        }
                    }
                }

                if (roomHasChanged)
                {
                    var roomchanges = EwsDatabase.SaveChanges();
                    Trace.WriteLine($"Rooms {roomchanges} saved to database.");
                }



                foreach (var room in EwsDatabase.RoomListRoomEntities.Where(w => !string.IsNullOrEmpty(w.Identity)))
                {
                    var roomservice = new EWService(Ewstoken);

                    EntitySubscription dbSubscription = null;
                    string watermark = null;
                    if (EwsDatabase.SubscriptionEntities.Any(rs => rs.SmtpAddress == room.SmtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.PullSubscription && !rs.Terminated))
                    {
                        dbSubscription = EwsDatabase.SubscriptionEntities.FirstOrDefault(fd => fd.SmtpAddress == room.SmtpAddress && fd.SubscriptionType == SubscriptionTypeEnum.PullSubscription && !fd.Terminated);
                        watermark = dbSubscription.Watermark;
                    }

                    var subscription = roomservice.CreatePullSubscription(ConnectingIdType.SmtpAddress, room.SmtpAddress, timeout, watermark);
                    if (!string.IsNullOrEmpty(watermark))
                    {
                        // close out the old subscription
                        dbSubscription.Terminated = true;
                    }

                    // newup a subscription to track the watermark
                    var newSubscription = new EntitySubscription()
                    {
                        Id = subscription.Id,
                        Watermark = subscription.Watermark,
                        LastRunTime = DateTime.UtcNow,
                        SubscriptionType = SubscriptionTypeEnum.PullSubscription,
                        SmtpAddress = room.SmtpAddress
                    };
                    EwsDatabase.SubscriptionEntities.Add(newSubscription);



                    subscriptions.Add(new SubscriptionCollection()
                    {
                        Runningsubscription = subscription,
                        SmtpAddress = room.SmtpAddress,
                        DatabaseSubscription = newSubscription
                    });
                }

                var rowChanged = EwsDatabase.SaveChanges();
                Trace.WriteLine($"Pull subscription persisted {rowChanged} rows");


                var waitTimer = new TimeSpan(0, 5, 0);
                while (!CancellationTokenSource.IsCancellationRequested)
                {
                    var milliseconds = (int)waitTimer.TotalMilliseconds;
                    PullRoomReservationChanges_Tick();
                    Trace.WriteLine($"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                    System.Threading.Thread.Sleep(milliseconds);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private void PullRoomReservationChanges_Tick()
        {
            // whatever you want to happen every 5 minutes
            Trace.WriteLine($"PullRoomReservationChangesAsync({mailboxOwner}) starting at {DateTime.UtcNow.ToShortTimeString()}");

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
                AppointmentSchema.Recurrence,
                EWSConstants.RefIdPropertyDef,
                EWSConstants.MeetingKeyPropertyDef);

            foreach (var item in subscriptions)
            {
                bool? ismore = default(bool);
                do
                {
                    PullSubscription subscription = item.Runningsubscription;
                    var appointmentMailbox = item.SmtpAddress;
                    var events = subscription.GetEvents();
                    var watermark = subscription.Watermark;
                    ismore = subscription.MoreEventsAvailable;

                    var dbroom = EwsDatabase.RoomListRoomEntities.FirstOrDefault(f => f.SmtpAddress == appointmentMailbox);

                    item.DatabaseSubscription.Watermark = watermark;
                    item.DatabaseSubscription.LastRunTime = DateTime.UtcNow;


                    // pull last event from stack TODO: need heuristic for how meetings can be stored
                    var filteredEvents = events.ItemEvents.Distinct(new ItemEventComparer()).OrderByDescending(x => x.TimeStamp);
                    foreach (ItemEvent ev in filteredEvents)
                    {
                        // Find an item event you can bind to
                        int action = 99;
                        var itemId = ev.ItemId;
                        int? meetingKey = default(int);
                        string refId = string.Empty;
                        string aptId = string.Empty;
                        var subscriptionitem = subscriptions.FirstOrDefault(k => k.Runningsubscription.Id == subscription.Id);
                        try
                        {
                            var appointmentTime = (Appointment)Item.Bind(subscription.Service, itemId, filterPropertySet);

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

                            // Initialize the calendar folder via Impersonation
                            EwsService.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxId);

                            try
                            {
                                CalendarFolder AtndCalendar = CalendarFolder.Bind(EwsService.Current, new FolderId(WellKnownFolderName.Calendar, mailboxId));
                                SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(CleanGlobalObjectId, Convert.ToBase64String((Byte[])CalIdVal));
                                ItemView ivItemView = new ItemView(5)
                                {
                                    PropertySet = new PropertySet(
                                        AppointmentSchema.Start,
                                        AppointmentSchema.End,
                                        AppointmentSchema.IsAllDayEvent,
                                        AppointmentSchema.IsRecurring,
                                        AppointmentSchema.IsCancelled,
                                        AppointmentSchema.TimeZone,
                                        EWSConstants.RefIdPropertyDef,
                                        EWSConstants.MeetingKeyPropertyDef)
                                };

                                FindItemsResults<Item> fiResults = AtndCalendar.FindItems(sfSearchFilter, ivItemView);
                                if (fiResults.Items.Count > 0)
                                {
                                    var filterApt = fiResults.Items.FirstOrDefault() as Appointment;
                                    Trace.WriteLine($"The first {fiResults.Items.Count()} appointments on your calendar from {filterApt.Start.ToShortDateString()} to {filterApt.End.ToShortDateString()}");



                                    var owernApptId = filterApt.Id;
                                    var ownerAppointmentTime = (Appointment)Item.Bind(EwsService.Current, owernApptId, filterPropertySet);


                                    var props = ownerAppointmentTime.ExtendedProperties.Where(p => (p.PropertyDefinition.PropertySet == DefaultExtendedPropertySet.Meeting));
                                    if (props.Any())
                                    {
                                        refId = (string)props.First(p => p.PropertyDefinition.Name == EWSConstants.RefIdPropertyName).Value;
                                        meetingKey = (int)props.First(p => p.PropertyDefinition.Name == EWSConstants.MeetingKeyPropertyName).Value;
                                    }

                                    // TODO: move this to the ServiceBus Processing
                                    if ((!string.IsNullOrEmpty(refId) || meetingKey.HasValue)
                                        && EwsDatabase.AppointmentEntities.Any(f => f.BookingReference == refId || f.Id == meetingKey))
                                    {
                                        var entity = EwsDatabase.AppointmentEntities.FirstOrDefault(f => f.BookingReference == refId || f.Id == meetingKey);

                                        entity.EndUTC = ownerAppointmentTime.End.ToUniversalTime();
                                        entity.StartUTC = ownerAppointmentTime.Start.ToUniversalTime();
                                        entity.ExistsInExchange = true;
                                        entity.IsRecurringMeeting = filterApt.IsRecurring;
                                        entity.Location = ownerAppointmentTime.Location;
                                        entity.OrganizerSmtpAddress = mailboxId;
                                        entity.Subject = ownerAppointmentTime.Subject;
                                        entity.RecurrencePattern = (ownerAppointmentTime.Recurrence == null) ? string.Empty : ownerAppointmentTime.Recurrence.ToString();
                                        entity.BookingReference = refId;
                                    }
                                    else
                                    {
                                        var entity = new EntityRoomAppointment()
                                        {
                                            BookingReference = refId,
                                            AppointmentUniqueId = itemId.UniqueId,
                                            EndUTC = ownerAppointmentTime.End.ToUniversalTime(),
                                            StartUTC = ownerAppointmentTime.Start.ToUniversalTime(),
                                            ExistsInExchange = true,
                                            IsRecurringMeeting = filterApt.IsRecurring,
                                            Location = ownerAppointmentTime.Location,
                                            OrganizerSmtpAddress = mailboxId,
                                            Subject = ownerAppointmentTime.Subject,
                                            RecurrencePattern = (ownerAppointmentTime.Recurrence == null) ? string.Empty : ownerAppointmentTime.Recurrence.ToString(),
                                            Room = dbroom
                                        };
                                        EwsDatabase.AppointmentEntities.Add(entity);
                                    }



                                    var task = System.Threading.Tasks.Task.Run(async () =>
                                    {
                                        await _queue.SendFromO365Async(appointmentMailbox, ownerAppointmentTime, action);
                                    });

                                    task.Wait(CancellationTokenSource.Token);
                                }


                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"Error retreiving calendar {mailboxId} msg:{ex.Message}");
                            }

                        }
                        catch (ServiceResponseException ex)
                        {
                            Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                            continue;
                        }
                    }

                }
                while (ismore == true);
            }

            var rowchanges = EwsDatabase.SaveChanges();
            Trace.WriteLine($"Saving subscription pull events with total rows {rowchanges}");
        }



        #region Trap application termination

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);

        private delegate bool EventHandler(CtrlType sig);
        static EventHandler _handler;

        enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        private static bool ConsoleCtrlCheck(CtrlType sig)
        {
            Trace.WriteLine("Exiting system due to external CTRL-C, or process kill, or shutdown");



            subscriptions.ToList().ForEach(subscription =>
            {
                var f = subscription.Runningsubscription;
                Trace.WriteLine($"Now unsubscribing for {f.Id}");
                f.Unsubscribe();

                subscription.DatabaseSubscription.Terminated = true;
                subscription.DatabaseSubscription.LastRunTime = DateTime.UtcNow;

            });

            var rowchanges = EwsDatabase.SaveChanges();
            Trace.WriteLine($"Saved subscription rows {rowchanges}");


            if (!IsDisposed)
            {
                EwsDatabase.Dispose();
            }

            CancellationTokenSource.Cancel();

            Trace.WriteLine("Cleanup complete");

            //allow main to run off
            exitSystem = true;

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }
        #endregion
    }
}
