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
using System.Data.Entity;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace EWSResourceSync
{
    class Program
    {
        static private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();

        static private bool IsDisposed { get; set; }

        static MessageManager Messenger { get; set; }

        static AuthenticationResult EwsToken { get; set; }

        private const int pollingTimeout = 30;


        #region Collections and Variables

        public Dictionary<string, GroupInfo> _groups { get; set; }

        public Mailboxes _mailboxes { get; set; }

        private Dictionary<string, StreamingSubscriptionConnection> _connections { get; set; }

        private List<SubscriptionCollection> _subscriptions { get; set; }


        private ITraceListener _traceListener { get; set; }

        private OAuthCredentials ServiceCredentials { get; set; }

        private bool _reconnect { get; set; }

        private Object _reconnectLock = new Object();

        private const int maxConcurrency = 100;

        /// <summary>
        /// Service Bus connection for sending subscription events
        /// </summary>
        private readonly string queueSubscription = EWSConstants.Config.ServiceBus.O365Subscription;

        /// <summary>
        /// Service Bus connection for sending change events from Folder Sync
        /// </summary>
        private readonly string queueSync = EWSConstants.Config.ServiceBus.O365Sync;

        #endregion



        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");


            var p = new Program();

            _handler += new EventHandler(p.ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            try
            {
                var service = System.Threading.Tasks.Task.Run(async () =>
                {
                    Trace.WriteLine("In Thread run await....");
                    await p.RunAsync();

                }, CancellationTokenSource.Token);
                service.Wait();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Failed in thread wait {ex.Message}");
            }
            finally
            {
                p.Dispose();
            }

            //hold the console so it doesn’t run off the end
            Trace.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }

        static Program()
        {
        }

        /// <summary>
        /// Establish the connection
        /// </summary>
        /// <returns></returns>
        async private System.Threading.Tasks.Task RunAsync()
        {
            _traceListener = new ClassTraceListener();

            IsDisposed = false;

            EwsToken = await EWSConstants.AcquireTokenAsync();

            Messenger = new MessageManager(CancellationTokenSource, EwsToken);

            ServiceCredentials = new OAuthCredentials(EwsToken.AccessToken);

            _groups = new Dictionary<string, GroupInfo>();
            _mailboxes = new Mailboxes(ServiceCredentials, _traceListener);
            _subscriptions = new List<SubscriptionCollection>();
            _connections = new Dictionary<string, StreamingSubscriptionConnection>();


            var impersonationId = EWSConstants.Config.Exchange.ImpersonationAcct;

            try
            {
                var list = new List<EWSFolderInfo>();
                using (EWSDbContext context = new EWSDbContext(EWSConstants.Config.Database))
                {
                    foreach (var room in context.RoomListRoomEntities.Where(w => !string.IsNullOrEmpty(w.Identity)))
                    {
                        var mailboxId = room.SmtpAddress;

                        try
                        {
                            var roomService = new EWService(EwsToken);
                            roomService.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxId);
                            var folderId = new FolderId(WellKnownFolderName.Calendar, mailboxId);
                            var info = new EWSFolderInfo()
                            {
                                SmtpAddress = mailboxId,
                                Service = roomService,
                                Folder = folderId
                            };
                            list.Add(info);
                        }
                        catch (Exception srex)
                        {
                            _traceListener.Trace("SyncProgram", $"Failed to ProcessChanges{srex.Message}");
                            throw new Exception($"ProcessChanges for {mailboxId} with MSG:{srex.Message}");
                        }
                    }
                }

                // Parallel ForEach (on RoomMailbox Grouping and SyncFolders) 
                await System.Threading.Tasks.Task.Run(() =>
                {
                    ParallelOptions options = new ParallelOptions
                    {
                        MaxDegreeOfParallelism = maxConcurrency
                    };

                    // Fireoff folder sync in background thread
                    Parallel.ForEach(list, options,
                        (bodyInfo) =>
                        {
                            ProcessChanges(bodyInfo);
                        });
                });


                var tasks = new List<System.Threading.Tasks.Task>();

                if (EWSConstants.Config.Exchange.PullEnabled)
                {
                    tasks.Add(PullSubscriptionChangesAsync(impersonationId));
                }

                // Upon completion kick of streamingsubscription
                tasks.Add(CreateStreamingSubscriptionGroupingAsync(impersonationId));

                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);
            }
        }

        /// <summary>
        /// Creates grouping subscriptions and waits for the notification events to flow
        /// </summary>
        /// <param name="impersonationId"></param>
        /// <returns></returns>
        async public System.Threading.Tasks.Task CreateStreamingSubscriptionGroupingAsync(string impersonationId)
        {
            var database = EWSConstants.Config.Database;

            using (EWSDbContext context = new EWSDbContext(database))
            {
                var smtpAddresses = await context.RoomListRoomEntities.ToListAsync();

                foreach (var sMailbox in smtpAddresses)
                {
                    var addedBox = _mailboxes.AddMailbox(sMailbox.SmtpAddress);
                    if (!addedBox)
                    {
                        _traceListener.Trace("SyncProgram", $"Failed to add SMTP {sMailbox.SmtpAddress}");
                    }

                    MailboxInfo mailboxInfo = _mailboxes.Mailbox(sMailbox.SmtpAddress);
                    if (mailboxInfo != null)
                    {
                        GroupInfo groupInfo = null;
                        if (_groups.ContainsKey(mailboxInfo.GroupName))
                        {
                            groupInfo = _groups[mailboxInfo.GroupName];
                        }
                        else
                        {
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, EwsToken, _traceListener);
                            _groups.Add(mailboxInfo.GroupName, groupInfo);
                        }

                        if (groupInfo.Mailboxes.Count > 199)
                        {
                            // We already have enough mailboxes in this group, so we rename it and create a new one
                            // Renaming it means that we can still put new mailboxes into the correct group based on GroupingInformation
                            int i = 1;
                            while (_groups.ContainsKey($"{groupInfo.Name}-{i}"))
                            {
                                i++;
                            }

                            // Remove previous grouping name from stack
                            _groups.Remove(groupInfo.Name);

                            // Add the grouping back with the new grouping name [keep the GroupInfo with previous name]
                            _groups.Add($"{groupInfo.Name}-{i}", groupInfo);

                            // Provision a new GroupInfo with the GroupName
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, EwsToken, _traceListener);

                            // Add GroupInfo to stack
                            _groups.Add(mailboxInfo.GroupName, groupInfo);
                        }

                        // Add the mailbox to the collection
                        groupInfo.Mailboxes.Add(sMailbox.SmtpAddress);
                    }
                }

                // Enable the Grouping
                foreach (var _group in _groups)
                {
                    _traceListener.Trace("SyncProgram", $"Opening connections for {_group.Key}");

                    var groupName = _group.Key;
                    var groupInfo = _group.Value;

                    // Create a streaming connection to the service object, over which events are returned to the client.
                    // Keep the streaming connection open for 30 minutes.
                    if (AddGroupSubscriptions(context, groupName))
                    {
                        _traceListener.Trace("SyncProgram", $"{groupInfo.Mailboxes.Count()} mailboxes primed for StreamingSubscriptions.");
                    }
                    else
                    {
                        _traceListener.Trace("SyncProgram", $"Group {groupInfo.Name} failed in StreamingSubscription events.");
                    }
                }

            }

            using (var semaphore = new System.Threading.SemaphoreSlim(1))
            {
                // Block the Thread
                semaphore.Wait();

                // Establish the StreamingSubscriptionConnections based on the GroupingInfo
                foreach (var connection in _connections)
                {
                    var _group = _groups[connection.Key];
                    var _connection = connection.Value;
                    if (_connection.IsOpen)
                    {
                        _group.IsConnectionOpen = true;
                        return;
                    }

                    try
                    {
                        _connection.Open();
                        _group.IsConnectionOpen = true;
                    }
                    catch (Exception ex)
                    {
                        _traceListener.Trace("SyncProgram", $"Error opening streamingsubscriptionconnection for group {_group.Name} MSG {ex.Message}");
                    }
                }

                // Lock the Thread Until its cancelled or fails
                await semaphore.WaitAsync(CancellationTokenSource.Token);
            }
        }

        /// <summary>
        /// Creates pull subscriptions for all rooms and waits on timer delay to pull subscriptions
        /// </summary>
        /// <param name="mailboxOwner"></param>
        /// <returns></returns>
        private async System.Threading.Tasks.Task PullSubscriptionChangesAsync(string mailboxOwner)
        {
            _traceListener.Trace("SyncProgram", $"PullSubscriptionChangesAsync({mailboxOwner}) starting");

            var service = new EWService(EwsToken);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            if (_subscriptions == null)
            {
                _subscriptions = new List<SubscriptionCollection>();
            }

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


                        _traceListener.Trace("SyncProgram", $"ListenToRoomReservationChangesAsync.Subscribed to room {room.SmtpAddress}");
                        _subscriptions.Add(new SubscriptionCollection()
                        {
                            Pulling = subscription,
                            SmtpAddress = room.SmtpAddress,
                            SubscriptionType = SubscriptionTypeEnum.PullSubscription
                        });

                        var rowChanged = _context.SaveChanges();
                        _traceListener.Trace("SyncProgram", $"Pull subscription persisted {rowChanged} rows");

                    }
                    catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
                    {
                        _traceListener.Trace("SyncProgram", $"Failed to provision subscription {srex.Message}");
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

                    using (var _context = new EWSDbContext(EWSConstants.Config.Database))
                    {

                        foreach (var item in _subscriptions)
                        {
                            bool? ismore = default(bool);
                            do
                            {
                                PullSubscription subscription = item.Pulling;
                                var events = subscription.GetEvents();
                                var watermark = subscription.Watermark;
                                ismore = subscription.MoreEventsAvailable;
                                var email = item.SmtpAddress;
                                var databaseItem = _context.SubscriptionEntities.FirstOrDefault(rs => rs.SmtpAddress == email && rs.SubscriptionType == SubscriptionTypeEnum.PullSubscription);

                                // pull last event from stack TODO: need heuristic for how meetings can be stored
                                var filteredEvents = events.ItemEvents.OrderBy(x => x.TimeStamp);
                                foreach (ItemEvent ev in filteredEvents)
                                {
                                    var itemId = ev.ItemId;
                                    try
                                    {
                                        // Send an item event you can bind to
                                        await Messenger.SendQueueO365ChangesAsync(queueSubscription, email, ev);
                                    }
                                    catch (ServiceResponseException ex)
                                    {
                                        _traceListener.Trace("SyncProgram", $"ServiceException: {ex.Message}");
                                        continue;
                                    }
                                }


                                databaseItem.Watermark = watermark;
                                databaseItem.LastRunTime = DateTime.UtcNow;

                                // Save Database changes
                                await _context.SaveChangesAsync();
                            }
                            while (ismore == true);
                        }
                    }

                    _traceListener.Trace("SyncProgram", $"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                    System.Threading.Thread.Sleep(milliseconds);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            _traceListener.Trace("SyncProgram", $"PullSubscriptionChangesAsync({mailboxOwner}) exiting");
        }

        /// <summary>
        /// Enable the synchronization of individual folders or Room[s]
        /// </summary>
        /// <param name="folderInfo"></param>
        public void ProcessChanges(EWSFolderInfo folderInfo)
        {
            bool moreChangesAvailable = false;
            _traceListener.Trace("SyncProgram", $"Entered ProcessChanges for {folderInfo.SmtpAddress} at {DateTime.UtcNow}");

            using (EWSDbContext context = new EWSDbContext(EWSConstants.Config.Database))
            {
                var room = context.RoomListRoomEntities.FirstOrDefault(f => f.SmtpAddress.Equals(folderInfo.SmtpAddress));
                var email = room.SmtpAddress;
                var syncState = room.SynchronizationState;
                var syncTimestamp = room.SynchronizationTimestamp;

                try
                {
                    do
                    {
                        // Get all changes since the last call. The synchronization cookie is stored in the _SynchronizationState field.
                        _traceListener.Trace("SyncProgram", $"Sync changes {email} with timestamp {syncTimestamp}");

                        // Just get the IDs of the items.
                        // For performance reasons, do not use the PropertySet.FirstClassProperties.
                        var changes = folderInfo.Service.Current.SyncFolderItems(folderInfo.Folder, PropertySet.IdOnly, null, 512, SyncFolderItemsScope.NormalItems, syncState);

                        // Update the synchronization 
                        syncState = changes.SyncState;

                        // Process all changes. If required, add a GetItem call here to request additional properties.
                        foreach (ItemChange itemChange in changes)
                        {
                            // This example just prints the ChangeType and ItemId to the console.
                            // A LOB application would apply business rules to each 
                            _traceListener.Trace("SyncProgram", $"ChangeType = {itemChange.ChangeType} with ItemId {itemChange.ItemId.ToString()}");

                            // Send an item event you can bind to
                            var service = System.Threading.Tasks.Task.Run(async () =>
                            {
                                _traceListener.Trace("SyncProgram", "In Thread run await....");
                                await Messenger.SendQueueO365SyncFoldersAsync(queueSync, email, itemChange);

                            }, CancellationTokenSource.Token);
                            service.Wait();
                        }

                        // If more changes are available, issue additional SyncFolderItems requests.
                        moreChangesAvailable = changes.MoreChangesAvailable;

                        room.SynchronizationState = syncState;
                        room.SynchronizationTimestamp = DateTime.UtcNow;

                        var roomchanges = context.SaveChanges();
                        _traceListener.Trace("SyncProgram", $"Room event folder sync {roomchanges} persisted.");
                    }
                    while (moreChangesAvailable);

                }
                catch (Exception processEx)
                {
                    Trace.TraceError($"Failed to send queue {processEx.Message}");
                }
                finally
                {
                    room.SynchronizationState = syncState;
                    room.SynchronizationTimestamp = DateTime.UtcNow;

                    var roomchanges = context.SaveChanges();
                    _traceListener.Trace("SyncProgram", $"Failed ProcessChanges({email}) folder sync {roomchanges} persisted.");
                }
            }
        }

        /// <summary>
        /// Adds streamingsubscription to the GroupInfo
        /// </summary>
        /// <param name="context">Database context</param>
        /// <param name="Group"></param>
        /// <param name="smtpAddress"></param>
        /// <returns></returns>
        private StreamingSubscription AddSubscription(EWSDbContext context, GroupInfo Group, string smtpAddress)
        {
            StreamingSubscription subscription = null;

            if (_subscriptions.Any(email => email.SmtpAddress.Equals(smtpAddress)))
            {
                var email = _subscriptions.FirstOrDefault(s => s.SmtpAddress.Equals(smtpAddress));
                _subscriptions.Remove(email);
            }

            try
            {
                ExchangeService exchange = Group.ExchangeService;
                exchange.Credentials = ServiceCredentials;
                exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);
                subscription = exchange.SubscribeToStreamingNotifications(new FolderId[] { WellKnownFolderName.Calendar },
                        EventType.Created,
                        EventType.Deleted,
                        EventType.Modified,
                        EventType.Moved,
                        EventType.Copied);


                _traceListener.Trace("SyncProgram", $"CreateStreamingSubscriptionGrouping to room {smtpAddress} with SubscriptionId {subscription.Id}");
                var subscriptionLastMark = default(DateTime?);
                var synchronizationState = string.Empty;

                EntitySubscription dbSubscription = null;
                if (context.SubscriptionEntities.Any(rs => rs.SmtpAddress == smtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription))
                {
                    dbSubscription = context.SubscriptionEntities.FirstOrDefault(rs => rs.SmtpAddress == smtpAddress && rs.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription);
                    subscriptionLastMark = dbSubscription.LastRunTime;
                    synchronizationState = dbSubscription.SynchronizationState;
                }
                else
                {   // newup a subscription to track the watermark
                    dbSubscription = new EntitySubscription()
                    {
                        LastRunTime = DateTime.UtcNow,
                        SubscriptionType = SubscriptionTypeEnum.StreamingSubscription,
                        SmtpAddress = smtpAddress
                    };
                    context.SubscriptionEntities.Add(dbSubscription);
                }

                dbSubscription.SubscriptionId = subscription.Id;

                var subscriptions = context.SaveChanges();
                _traceListener.Trace("SyncProgram", $"Streaming subscription persisted {subscriptions} persisted.");


                _subscriptions.Add(new SubscriptionCollection()
                {
                    SmtpAddress = smtpAddress,
                    Streaming = subscription,
                    SubscriptionType = SubscriptionTypeEnum.StreamingSubscription,
                    SynchronizationDateTime = subscriptionLastMark,
                    SynchronizationState = synchronizationState
                });

            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
            {
                _traceListener.Trace("SyncProgram", $"Failed to provision subscription {srex.Message}");
                throw new Exception($"Subscription could not be created for {smtpAddress} with MSG:{srex.Message}");
            }

            return subscription;
        }

        /// <summary>
        /// Process the GroupInfo and create subscriptions based on AnchorMailbox
        /// </summary>
        /// <param name="context">Database context for subscription update</param>
        /// <param name="groupName">EWS GroupInfo or Dynamic Groupname if Mailboxes > 200</param>
        /// <returns></returns>
        public bool AddGroupSubscriptions(EWSDbContext context, string groupName)
        {
            if (!_groups.ContainsKey(groupName))
                return false;

            if (_connections == null)
            {
                _connections = new Dictionary<string, StreamingSubscriptionConnection>();
            }


            if (_connections.ContainsKey(groupName))
            {
                var _connection = _connections[groupName];

                foreach (StreamingSubscription subscription in _connection.CurrentSubscriptions)
                {
                    try
                    {
                        subscription.Unsubscribe();
                    }
                    catch { }
                }
                try
                {
                    _connection.Close();

                    _groups[groupName].IsConnectionOpen = false;
                }
                catch { }
            }

            try
            {
                // Create the connection for this group, and the primary mailbox subscription
                var groupInfo = _groups[groupName];
                var PrimaryMailbox = groupInfo.PrimaryMailbox;
                // Return the subscription, or create a new one if we don't already have one
                StreamingSubscription mailboxSubscription = AddSubscription(context, groupInfo, PrimaryMailbox);
                if (_connections.ContainsKey(groupName))
                {
                    _connections[groupName] = new StreamingSubscriptionConnection(mailboxSubscription.Service, pollingTimeout);
                }
                else
                {
                    _connections.Add(groupName, new StreamingSubscriptionConnection(mailboxSubscription.Service, pollingTimeout));
                }


                SubscribeConnectionEvents(_connections[groupName]);
                _connections[groupName].AddSubscription(mailboxSubscription);
                _traceListener.Trace("SyncProgram", $"{PrimaryMailbox} (primary mailbox) subscription created in group {groupName}");

                // Now add any further subscriptions in this group
                foreach (string sMailbox in groupInfo.Mailboxes.Where(w => !w.Equals(PrimaryMailbox)))
                {
                    try
                    {
                        var localSubscription = AddSubscription(context, groupInfo, sMailbox);
                        _connections[groupName].AddSubscription(localSubscription);
                        _traceListener.Trace("Add secondary subscription", $"{sMailbox} subscription created in group {groupName}");
                    }
                    catch (Exception ex)
                    {
                        _traceListener.Trace("Exception", $"ERROR when subscribing {sMailbox} in group {groupName}: {ex.Message}");
                    }

                }
            }
            catch (Exception ex)
            {
                _traceListener.Trace("Exception", $"ERROR when creating subscription connection group {groupName}: {ex.Message}");
            }
            return true;
        }

        /// <summary>
        /// Subscribe to events for this connection
        /// </summary>
        /// <param name="connection"></param>
        private void SubscribeConnectionEvents(StreamingSubscriptionConnection connection)
        {
            connection.OnNotificationEvent += Connection_OnNotificationEvent;
            connection.OnDisconnect += Connection_OnDisconnect;
            connection.OnSubscriptionError += Connection_OnSubscriptionError;
        }

        /// <summary>
        /// Write Subscription Errors TODO: Increase error handling
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Connection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                _traceListener.Trace("SyncProgram", $"OnSubscriptionError received for {args.Subscription.Service.ImpersonatedUserId.Id}.");
                _traceListener.Trace("SyncProgram", $"OnSubscriptionError(Exception) : {args.Exception.Message} Stack Trace : {args.Exception.StackTrace} Inner Exception : {args.Exception.InnerException}");
            }
            catch
            {
                _traceListener.Trace("SyncProgram", "OnSubscriptionError received");
            }
        }

        /// <summary>
        /// Disconnected subscription
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Connection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                _traceListener.Trace("SyncProgram", $"StreamingSubscriptionChangesAsync OnDisconnect with exception: {args.Exception}");
                _traceListener.Trace("SyncProgram", $"OnDisconnection received for {args.Subscription.Service.ImpersonatedUserId.Id}");

                if (CancellationTokenSource.IsCancellationRequested)
                {
                    _traceListener.Trace("SyncProgram", $"OnDisconnect Closing streamingsubscriptionconnection at {DateTime.UtcNow}..");
                    CloseConnections();
                }
                else
                {
                    _traceListener.Trace("SyncProgram", $"StreamingSubscriptionChangesAsync re-connect");

                    if (!_reconnect)
                        return;

                    ReconnectToSubscriptions();
                }
            }
            catch
            {
                _traceListener.Trace("SyncProgram", "OnDisconnection received");
            }
            _reconnect = true;  // We can't reconnect in the disconnect event, so we set a flag for the timer to pick this up and check all the connections
        }

        /// <summary>
        /// Process individual subscription notifications
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Connection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            var ewsService = args.Subscription.Service;
            var fromEmailAddress = args.Subscription.Service.ImpersonatedUserId.Id;
            _traceListener.Trace("SyncProgram", $"StreamingSubscriptionChangesAsync {fromEmailAddress} received {args.Events.Count()} notification(s)");

            // Process all Events
            // New appt: ev.Type == Created
            // Del apt:  ev.Type == Moved, IsCancelled == true
            // Move apt: ev.Type == Modified, IsUnmodified == false, 
            foreach (NotificationEvent notification in args.Events)
            {

                if (string.IsNullOrEmpty(fromEmailAddress))
                    fromEmailAddress = "Unknown mailbox";

                string sEvent = fromEmailAddress + ": ";

                if (notification is ItemEvent)
                {
                    ItemEvent ev = (ItemEvent)notification;

                    var itemId = ev.ItemId;
                    sEvent += $"Item {ev.EventType.ToString()}: ";
                    sEvent += $"ItemId = {ev.ItemId.UniqueId}";
                    _traceListener.Trace("SyncProgram", $"Processing event {sEvent} at {DateTime.UtcNow}");

                    try
                    {
                        var task = System.Threading.Tasks.Task.Run(async () =>
                        {
                            // Send an item event you can bind to
                            await Messenger.SendQueueO365ChangesAsync(queueSubscription, fromEmailAddress, ev);
                        });

                        task.Wait(CancellationTokenSource.Token);
                    }
                    catch (ServiceResponseException ex)
                    {
                        _traceListener.Trace("SyncProgram", $"ServiceException: {ex.Message}");
                        continue;
                    }
                }
                else if (notification is FolderEvent)
                {
                    FolderEvent ev = (FolderEvent)notification;

                    sEvent += $"Folder {ev.EventType.ToString()}: ";
                    sEvent += $"FolderId = {ev.FolderId.UniqueId}";
                    _traceListener.Trace("SyncProgram", $"Processing event {sEvent} at {DateTime.UtcNow}");
                }
            }
        }

        /// <summary>
        /// Reconnect StreamingSubscriptionConnections [if not in a disconnecting or error'd state]
        /// </summary>
        public void ReconnectToSubscriptions()
        {
            // Go through our connections and reconnect any that have closed
            _reconnect = false;
            lock (_reconnectLock)  // Prevent this code being run concurrently (i.e. if an event fires in the middle of the processing)
            {
                using (EWSDbContext context = new EWSDbContext(EWSConstants.Config.Database))
                {
                    foreach (var _connectionPair in _connections)
                    {
                        var connectionGroupName = _connectionPair.Key;
                        var groupInfo = _groups[connectionGroupName];
                        var _connection = _connections[connectionGroupName];
                        if (!_connection.IsOpen)
                        {
                            try
                            {
                                try
                                {
                                    _connection.Open();
                                    _traceListener.Trace("SyncProgram", $"Re-opened connection for group {connectionGroupName}");
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message.StartsWith("You must add at least one subscription to this connection before it can be opened"))
                                    {
                                        // Try recreating this group
                                        AddGroupSubscriptions(context, connectionGroupName);
                                    }
                                    else
                                    {
                                        _traceListener.Trace("SyncProgram", $"Failed to reopen connection: {ex.Message}");
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                _traceListener.Trace("SyncProgram", $"Failed to reopen connection: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Close connections and unsubscribe
        /// </summary>
        public void CloseConnections()
        {
            _traceListener.Trace("SyncProgram", $"ConsoleCtrlCheck CloseConnections at {DateTime.UtcNow}..");

            foreach (StreamingSubscriptionConnection connection in _connections.Values)
            {
                if (connection.IsOpen)
                {
                    connection.Close();
                }
            }

            // Unsubscribe all
            if (_subscriptions == null)
                return;


            using (EWSDbContext context = new EWSDbContext(EWSConstants.Config.Database))
            {
                var databaseItems = context.SubscriptionEntities.Where(di => di.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription);

                try
                {
                    for (int i = _subscriptions.Count - 1; i >= 0; i--)
                    {
                        SubscriptionCollection subscriptionItem = _subscriptions[i];
                        var subscription = subscriptionItem.Streaming;
                        try
                        {
                            subscription.Unsubscribe();
                            _traceListener.Trace("SyncProgram", $"Unsubscribed from {subscriptionItem.SmtpAddress}");

                            if (databaseItems.Any(di => di.SmtpAddress == subscriptionItem.SmtpAddress && di.SubscriptionType == SubscriptionTypeEnum.StreamingSubscription))
                            {
                                var item = databaseItems.FirstOrDefault();
                                item.Terminated = true;
                                item.LastRunTime = DateTime.UtcNow;
                            }
                        }
                        catch (Exception ex)
                        {
                            _traceListener.Trace("SyncProgram", $"Error when unsubscribing {subscriptionItem.SmtpAddress}: {ex.Message}");
                        }

                        _subscriptions.Remove(subscriptionItem);
                    }
                }
                catch (Exception subEx)
                {
                    _traceListener.Trace("SyncProgram", $"Failed in subscription disconnect {subEx.Message}");
                }
                finally
                {
                    var changes = context.SaveChanges();
                    _traceListener.Trace("SyncProgram", $"Subscription watermark sync changed {changes} rows.");
                }
            }

            _reconnect = false;
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

        private bool ConsoleCtrlCheck(CtrlType sig)
        {
            _traceListener.Trace("SyncProgram", "Exiting system due to external CTRL-C, or process kill, or shutdown");

            // dispose all threads
            Dispose();

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }

        private void Dispose()
        {
            Trace.WriteLine("Disposing");
            if (IsDisposed)
                return;

            // should cancel all registered events
            CancellationTokenSource.Cancel();

            // issue into messenger
            Messenger.IssueCancellation(CancellationTokenSource);

            // Close the connection stream
            CloseConnections();

            // should close out database and issue cancellation to token
            Messenger.Dispose();

            IsDisposed = true;
            _traceListener.Trace("SyncProgram", "Cleanup complete");
        }

        #endregion
    }
}
