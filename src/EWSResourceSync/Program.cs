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
using System.Threading;

namespace EWSResourceSync
{
    class Program
    {
        static private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        static private bool IsDisposed { get; set; }
        static MessageManager Messenger { get; set; }
        static AuthenticationResult EwsToken { get; set; }
        private static int pollingTimeout = 30;

        private static List<StreamingSubscriptionConnection> _connections = new List<StreamingSubscriptionConnection>();
        private static List<SubscriptionCollection> subscriptions { get; set; }

        private static StreamingSubscriptionConnection connection { get; set; }


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

            EwsToken = await EWSConstants.AcquireTokenAsync();

            Messenger = new MessageManager(CancellationTokenSource, EwsToken);

            var queueSubscription = EWSConstants.Config.ServiceBus.O365Subscription;
            var impersonationId = EWSConstants.Config.Exchange.ImpersonationAcct;

            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                if (EWSConstants.Config.Exchange.PullEnabled)
                {
                    tasks.Add(PullSubscriptionChangesAsync(queueSubscription, impersonationId));
                }

                tasks.Add(StreamingSubscriptionChangesAsync(queueSubscription, impersonationId));

                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }



        public List<SubscriptionCollection> CreateStreamingSubscriptionGrouping()
        {
            var subscriptions = new List<SubscriptionCollection>();
            var EwsService = new EWService(Ewstoken);

            EwsService.SetImpersonation(ConnectingIdType.SmtpAddress, MailboxOwner);

            foreach (var room in EwsDatabase.RoomListRoomEntities.Where(w => !string.IsNullOrEmpty(w.Identity)))
            {
                var mailboxId = room.SmtpAddress;

                EntitySubscription dbSubscription = null;
                var subscriptionLastMark = default(DateTime?);
                var synchronizationState = string.Empty;
                if (EwsDatabase.SubscriptionEntities.Any(rs => rs.SmtpAddress == mailboxId && rs.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription))
                {
                    dbSubscription = EwsDatabase.SubscriptionEntities.FirstOrDefault(rs => rs.SmtpAddress == room.SmtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription);
                    subscriptionLastMark = dbSubscription.LastRunTime;
                    synchronizationState = dbSubscription.SynchronizationState;
                }
                else
                {   // newup a subscription to track the watermark
                    dbSubscription = new EntitySubscription()
                    {
                        LastRunTime = DateTime.UtcNow,
                        SubscriptionType = SubscriptionTypeEnum.StreamingSubscription,
                        SmtpAddress = mailboxId
                    };
                    EwsDatabase.SubscriptionEntities.Add(dbSubscription);
                }

                try
                {
                    var roomService = new EWService(Ewstoken);
                    roomService.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxId);
                    var folderId = new FolderId(WellKnownFolderName.Calendar, mailboxId);
                    var info = new EWSFolderInfo()
                    {
                        SmtpAddress = mailboxId,
                        SynchronizationState = synchronizationState,
                        Service = roomService,
                        Folder = folderId,
                        LastRunTime = subscriptionLastMark
                    };

                    // Fireoff folder sync in background thread
                    ThreadPool.QueueUserWorkItem(new WaitCallback(ProcessChanges), info);
                }
                catch (Exception srex)
                {
                    Trace.WriteLine($"Failed to ProcessChanges{srex.Message}");
                    throw new Exception($"ProcessChanges for {mailboxId} with MSG:{srex.Message}");
                }

                try
                {
                    var roomService = new EWService(Ewstoken);
                    var subscription = roomService.CreateStreamingSubscription(ConnectingIdType.SmtpAddress, mailboxId);



                    Trace.WriteLine($"CreateStreamingSubscriptionGrouping to room {mailboxId}");
                    subscriptions.Add(new SubscriptionCollection()
                    {
                        Streaming = subscription,
                        SmtpAddress = mailboxId,
                        DatabaseSubscription = dbSubscription,
                        SubscriptionType = SubscriptionTypeEnum.StreamingSubscription
                    });

                }
                catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
                {
                    Trace.WriteLine($"Failed to provision subscription {srex.Message}");
                    throw new Exception($"Subscription could not be created for {mailboxId} with MSG:{srex.Message}");
                }
            }

            try
            {
                var rowChanged = EwsDatabase.SaveChanges();
                Trace.WriteLine($"Streaming subscription persisted {rowChanged} rows");
            }
            catch (System.Data.Entity.Core.EntityException dex)
            {
                Trace.WriteLine($"Failed to EF persist {dex.Message}");
            }

            return subscriptions;
        }



        private async System.Threading.Tasks.Task StreamingSubscriptionChangesAsync(string queueConnection, string mailboxOwner)
        {
            Trace.WriteLine($"StreamingSubscriptionChangesAsync({mailboxOwner}) starting");

            var service = new EWService(EwsToken);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            subscriptions = Messenger.CreateStreamingSubscriptionGrouping();

            // Create a streaming connection to the service object, over which events are returned to the client.
            // Keep the streaming connection open for 30 minutes.
            connection = new StreamingSubscriptionConnection(service.Current, subscriptions.Select(s => s.Streaming), 30);
            var semaphore = new System.Threading.SemaphoreSlim(1);

            connection.OnNotificationEvent += (object sender, NotificationEventArgs args) =>
            {
                string fromEmailAddress = args.Subscription.Service.ImpersonatedUserId.Id;
                Trace.WriteLine($"StreamingSubscriptionChangesAsync received {args.Events.Count()} notification(s)");


                var ewsService = args.Subscription.Service;
                var subscription = subscriptions.FirstOrDefault(k => k.Streaming.Id == args.Subscription.Id);
                var watermark = subscription.Streaming.Watermark;
                var databaseItem = subscription.DatabaseSubscription;
                var email = subscription.SmtpAddress;

                // Process all Events
                // New appt: ev.Type == Created
                // Del apt:  ev.Type == Moved, IsCancelled == true
                // Move apt: ev.Type == Modified, IsUnmodified == false, 
                var evGroups = args.Events.Where(ev => ev is ItemEvent).Select(ev => ((ItemEvent)ev)).OrderBy(x => x.TimeStamp);
                foreach (ItemEvent ev in evGroups)
                {
                    var itemId = ev.ItemId;
                    try
                    {
                        var task = System.Threading.Tasks.Task.Run(async () =>
                        {
                            databaseItem.Watermark = watermark;
                            databaseItem.LastRunTime = DateTime.UtcNow;

                            // Send an item event you can bind to
                            await Messenger.SendQueueO365ChangesAsync(queueConnection, email, ev);

                            await Messenger.SaveDbChangesAsync();
                        });

                        task.Wait(CancellationTokenSource.Token);
                    }
                    catch (ServiceResponseException ex)
                    {
                        Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                        continue;
                    }
                }

            };
            connection.OnDisconnect += (object sender, SubscriptionErrorEventArgs args) =>
            {
                Trace.WriteLine($"StreamingSubscriptionChangesAsync OnDisconnect with exception: {args.Exception}");

                if (CancellationTokenSource.IsCancellationRequested)
                {
                    Trace.WriteLine($"StreamingSubscriptionChangesAsync disconnecting");
                    if (connection.CurrentSubscriptions != null && subscriptions != null)
                    {
                        Trace.WriteLine($"OnDisconnect Closing streamingsubscriptionconnection at {DateTime.UtcNow}..");
                        foreach (var item in subscriptions)
                        {
                            item.DatabaseSubscription.LastRunTime = DateTime.UtcNow;
                            item.DatabaseSubscription.Terminated = true;

                            Trace.WriteLine($"RemoveSubscription {item.Streaming.Id}");
                            connection.RemoveSubscription(item.Streaming);
                        };
                    }
                    semaphore.Release();
                }
                else
                {
                    Trace.WriteLine($"StreamingSubscriptionChangesAsync re-connected");
                    connection.Open();
                }
            };
            connection.OnSubscriptionError += (object sender, SubscriptionErrorEventArgs args) =>
            {
                if (args.Exception != null)
                {
                    Trace.WriteLine($"OnSubscriptionError(Exception) : {args.Exception.Message} Stack Trace : {args.Exception.StackTrace} Inner Exception : {args.Exception.InnerException}");
                }
            };

            semaphore.Wait();
            connection.Open();
            await semaphore.WaitAsync(CancellationTokenSource.Token);

            Trace.WriteLine($"StreamingSubscriptionChangesAsync({mailboxOwner}) exiting");
        }

        private async System.Threading.Tasks.Task PullSubscriptionChangesAsync(string queueConnection, string mailboxOwner)
        {
            Trace.WriteLine($"PullSubscriptionChangesAsync({mailboxOwner}) starting");

            var service = new EWService(EwsToken);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);


            subscriptions = new List<SubscriptionCollection>();
            var EwsService = new EWService(EwsToken);
            EwsService.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            // Retreive and Store PullSubscription Details
            using (var _context = new EWSDbContext(EWSConstants.Config.Database))
            {

                foreach (var room in _context.RoomListRoomEntities.Where(w => !string.IsNullOrEmpty(w.Identity)))
                {
                    EntitySubscription dbSubscription = null;
                    string watermark = null;
                    if (_context.SubscriptionEntities.Any(rs => rs.SmtpAddress == room.SmtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.PullSubscription))
                    {
                        dbSubscription = _context.SubscriptionEntities.FirstOrDefault(rs => rs.SmtpAddress == room.SmtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.PullSubscription);
                        watermark = dbSubscription.Watermark;
                    }
                    else
                    {
                        // newup a subscription to track the watermark
                        dbSubscription = new EntitySubscription()
                        {
                            LastRunTime = DateTime.UtcNow,
                            SubscriptionType = SubscriptionTypeEnum.PullSubscription,
                            SmtpAddress = room.SmtpAddress
                        };
                        _context.SubscriptionEntities.Add(dbSubscription);
                    }

                    try
                    {
                        var roomService = new EWService(EwsToken);
                        var subscription = roomService.CreatePullSubscription(ConnectingIdType.SmtpAddress, room.SmtpAddress, pollingTimeout, watermark);

                        // close out the old subscription
                        dbSubscription.PreviousWatermark = (!string.IsNullOrEmpty(watermark)) ? watermark : null;
                        dbSubscription.SubscriptionId = subscription.Id;
                        dbSubscription.Watermark = subscription.Watermark;


                        Trace.WriteLine($"ListenToRoomReservationChangesAsync.Subscribed to room {room.SmtpAddress}");
                        subscriptions.Add(new SubscriptionCollection()
                        {
                            Pulling = subscription,
                            SmtpAddress = room.SmtpAddress,
                            DatabaseSubscription = dbSubscription,
                            SubscriptionType = SubscriptionTypeEnum.PullSubscription
                        });

                        var rowChanged = _context.SaveChanges();
                        Trace.WriteLine($"Pull subscription persisted {rowChanged} rows");

                    }
                    catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
                    {
                        Trace.WriteLine($"Failed to provision subscription {srex.Message}");
                        throw new Exception($"Subscription could not be created for {room.SmtpAddress} with MSG:{srex.Message}");
                    }
                }

            }


            try
            {
                var waitTimer = new TimeSpan(0, 5, 0);
                while (!CancellationTokenSource.IsCancellationRequested)
                {
                    var milliseconds = (int)waitTimer.TotalMilliseconds;

                    foreach (var item in subscriptions)
                    {
                        bool? ismore = default(bool);
                        do
                        {
                            PullSubscription subscription = item.Pulling;
                            var events = subscription.GetEvents();
                            var watermark = subscription.Watermark;
                            ismore = subscription.MoreEventsAvailable;
                            var databaseItem = item.DatabaseSubscription;
                            var email = item.SmtpAddress;


                            // pull last event from stack TODO: need heuristic for how meetings can be stored
                            var filteredEvents = events.ItemEvents.OrderBy(x => x.TimeStamp);
                            foreach (ItemEvent ev in filteredEvents)
                            {
                                var itemId = ev.ItemId;
                                try
                                {
                                    databaseItem.Watermark = watermark;
                                    databaseItem.LastRunTime = DateTime.UtcNow;

                                    // Send an item event you can bind to
                                    await Messenger.SendQueueO365ChangesAsync(queueConnection, email, ev);
                                    // Save Database changes
                                    await Messenger.SaveDbChangesAsync();
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

                    Trace.WriteLine($"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                    System.Threading.Thread.Sleep(milliseconds);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Trace.WriteLine($"PullSubscriptionChangesAsync({mailboxOwner}) exiting");
        }

        /// <summary>
        /// Enable the synchronization of individual folders or Room[s]
        /// </summary>
        /// <param name="folderInfo"></param>
        public static void ProcessChanges(object folderInfo)
        {
            bool moreChangesAvailable;
            EWSFolderInfo info = (EWSFolderInfo)folderInfo;
            do
            {
                // Get all changes since the last call. The synchronization cookie is stored in the _SynchronizationState field.
                // Just get the IDs of the items.
                // For performance reasons, do not use the PropertySet.FirstClassProperties.
                var changes = info.Service.Current.SyncFolderItems(info.Folder, PropertySet.IdOnly, null, 512, SyncFolderItemsScope.NormalItems, info.SynchronizationState);

                // Update the synchronization 
                info.SynchronizationState = changes.SyncState;

                // Process all changes. If required, add a GetItem call here to request additional properties.
                foreach (ItemChange itemChange in changes)
                {
                    // This example just prints the ChangeType and ItemId to the console.
                    // A LOB application would apply business rules to each 
                    Trace.WriteLine($"ChangeType = {itemChange.ChangeType} with ItemId {itemChange.ItemId.ToString()}");
                }

                // If more changes are available, issue additional SyncFolderItems requests.
                moreChangesAvailable = changes.MoreChangesAvailable;
            }
            while (moreChangesAvailable);
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


            // should cancel all registered events
            CancellationTokenSource.Cancel();

            // issue into messenger
            Messenger.IssueCancellation(CancellationTokenSource);

            // Close the connection stream
            if (connection != null && connection.IsOpen)
            {
                Trace.WriteLine($"ConsoleCtrlCheck Closing streamingsubscriptionconnection at {DateTime.UtcNow}..");
                connection.Close();
            }

            // dispose of the messenger
            if (!IsDisposed)
            {
                // should close out database and issue cancellation to token
                Messenger.Dispose();
            }

            // cleanup complete
            Trace.WriteLine("Cleanup complete");

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }

        #endregion
    }
}
