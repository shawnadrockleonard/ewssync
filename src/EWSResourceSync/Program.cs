using EWS.Common;
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

namespace EWSResourceSync
{
    class Program
    {
        static private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        static private bool exitSystem = false;
        static private bool IsDisposed { get; set; }
        static MessageManager Messenger { get; set; }

        private static List<StreamingSubscriptionConnection> _connections = new List<StreamingSubscriptionConnection>();

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");

            _handler += new EventHandler(ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            var p = new Program();

            var service = System.Threading.Tasks.Task.Run(async () =>
            {
                Trace.WriteLine("In Thread run await....");
                await p.RunAsync();

            }, CancellationTokenSource.Token);
            service.Wait();

            //hold the console so it doesn’t run off the end
            Console.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }

        static Program()
        {
        }

        async private System.Threading.Tasks.Task RunAsync()
        {
            IsDisposed = false;

            var Ewstoken = await EWSConstants.AcquireTokenAsync();

            Messenger = new MessageManager(CancellationTokenSource, Ewstoken);

            var queueConnection = EWSConstants.Config.ServiceBus.O365Subscription;


            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                tasks.Add(Messenger.ReceiveQueueO365ChangesAsync(queueConnection));
                tasks.Add(ListenToRoomReservationChangesAsync(EWSConstants.Config.Exchange.ImpersonationAcct));

                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private Dictionary<StreamingSubscription, string> subscriptions;

        // # TODO: Need to handle subscription for re-hydration
        // # TODO: evaluate why extension attributes are not returned in listener

        private async System.Threading.Tasks.Task ListenToRoomReservationChangesAsync(string mailboxOwner)
        {
            Trace.WriteLine($"ListenToRoomReservationChangesAsync({mailboxOwner}) starting");

            
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            subscriptions = new Dictionary<StreamingSubscription, string>();

            var roomlisting = service.GetRoomListing();
            foreach (var roomlist in roomlisting)
            {
                foreach (var room in roomlist.Value)
                {
                    try
                    {
                        var roomService = new EWService(tokens);
                        var sub = roomService.CreateStreamingSubscription(ConnectingIdType.SmtpAddress, room.Address);

                        Trace.WriteLine($"ListenToRoomReservationChangesAsync.Subscribed to room {room.Address}");
                        subscriptions.Add(sub, room.Address);
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
            var connection = new StreamingSubscriptionConnection(service.Current, subscriptions.Keys.Select(s => s), 30);
            var semaphore = new System.Threading.SemaphoreSlim(1);

            connection.OnNotificationEvent += OnNotificationEvent;
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


                Trace.WriteLine($"ListenToRoomReservationChangesAsync received notification");


                ExtendedPropertyDefinition CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);

                // New appt: ev.Type == Created
                // Del apt:  ev.Type == Moved, IsCancelled == true
                // Move apt: ev.Type == Modified, IsUnmodified == false, 
                var evGroups = args.Events.Where(ev => ev is ItemEvent).Select(ev => ((ItemEvent)ev));

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
                        EWSConstants.RefIdPropertyDef,
                        EWSConstants.MeetingKeyPropertyDef);

                var polingService = new EWService(tokens);


                // TODO: Application centric notification stream (Needs to handle all of the Events - process in linear fashion)
                // Process all Events

                foreach (ItemEvent ev in evGroups)
                {
                    var itemId = ev.ItemId;

                    // Find an item event you can bind to
                    var ewsService = args.Subscription.Service;
                    var subscription = subscriptions.FirstOrDefault(k => k.Key.Id == args.Subscription.Id);

                    try
                    {
                        ewsService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, subscription.Value);

                        var appointmentTime = service.GetAppointment(ewsService, itemId, filterPropertySet);

                        var parentAppointment = service.GetParentAppointment(ewsService, appointmentTime, filterPropertySet);

                        await _queue.SendFromO365Async(ewsService.ImpersonatedUserId.Id, parentAppointment, ev.EventType);
                    }
                    catch (ServiceResponseException ex)
                    {
                        Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                        continue;
                    }


                }
            
        }



        void ShowMoreInfo(object e)
        {
            // Get more info for the given item.  This will run on it's own thread
            // so that the main program can continue as usual (we won't hold anything up)

            NotificationInfo n = (NotificationInfo)e;

            var service = new EWService(tokens, true);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, n.Mailbox);

            service.Current.Url = n.Service.Url;
            service.Current.TraceFlags = TraceFlags.All;

            var ewsMoreInfoService = service.Current;


            string sEvent = "";
            if (n.Event is ItemEvent)
            {
                sEvent = n.Mailbox + ": Item " + (n.Event as ItemEvent).EventType.ToString() + ": " + MoreItemInfo(n.Event as ItemEvent, ewsMoreInfoService);
            }
            else
                sEvent = n.Mailbox + ": Folder " + (n.Event as FolderEvent).EventType.ToString() + ": " + MoreFolderInfo(n.Event as FolderEvent, ewsMoreInfoService);

            ShowEvent(sEvent);
        }

        private void ShowEvent(string eventDetails)
        {
            try
            {

            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Ex {ex.Message} Error");
            }
        }

        private string MoreItemInfo(ItemEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Item subject=" + GetItemInfo(e.ItemId, service);

            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        private string MoreFolderInfo(FolderEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Folder name=" + GetFolderName(e.FolderId, service);
            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        private string GetItemInfo(ItemId itemId, ExchangeService service)
        {
            // Retrieve the subject for a given item
            string sItemInfo = "";
            Item oItem;
            PropertySet oPropertySet;


            oPropertySet = new PropertySet(ItemSchema.Subject);

            try
            {
                oItem = Item.Bind(service, itemId, oPropertySet);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            if (oItem is Appointment)
            {
                sItemInfo += "Appointment subject=" + oItem.Subject;
                // Show attendee information
                Appointment oAppt = oItem as Appointment;
                sItemInfo += ",RequiredAttendees=" + GetAttendees(oAppt.RequiredAttendees);
                sItemInfo += ",OptionalAttendees=" + GetAttendees(oAppt.OptionalAttendees);
            }
            else
                sItemInfo += "Item subject=" + oItem.Subject;


            return sItemInfo;
        }

        private string GetAttendees(AttendeeCollection attendees)
        {
            if (attendees.Count == 0) return "none";

            string sAttendees = "";
            foreach (Attendee attendee in attendees)
            {
                if (!String.IsNullOrEmpty(sAttendees))
                    sAttendees += ", ";
                sAttendees += attendee.Name;
            }

            return sAttendees;
        }

        private string GetFolderName(FolderId folderId, ExchangeService service)
        {
            // Retrieve display name of the given folder
            try
            {
                Folder oFolder = Folder.Bind(service, folderId, new PropertySet(FolderSchema.DisplayName));
                return oFolder.DisplayName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
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
