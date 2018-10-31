using EWS.Common.Database;
using EWS.Common.Models;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class GroupInfoSubscriptionManager : IDisposable
    {
        public GroupInfoSubscriptionManager(ITraceListener listener)
        {
            _traceListener = listener;

            var EwsToken = System.Threading.Tasks.Task.Run(async () =>
            {
                return await EWSConstants.AcquireTokenAsync();
            });

            ewsToken = EwsToken.Result;

            credentials = new OAuthCredentials(EwsToken.Result.AccessToken);

            _groups = new Dictionary<string, GroupInfo>();
            _mailboxes = new Mailboxes(credentials, _traceListener);
            _subscriptions = new List<SubscriptionCollection>();
            _connections = new Dictionary<string, StreamingSubscriptionConnection>();
        }

        #region Collections and Variables

        public Dictionary<string, GroupInfo> _groups { get; set; }

        public Mailboxes _mailboxes { get; set; }

        private Dictionary<string, StreamingSubscriptionConnection> _connections { get; set; }

        private List<SubscriptionCollection> _subscriptions { get; set; }


        private ITraceListener _traceListener { get; set; }

        private AuthenticationResult ewsToken { get; set; }

        private OAuthCredentials credentials { get; set; }

        private bool _reconnect { get; set; }

        private Object _reconnectLock = new Object();

        private const int numericUpDownTimeout = 30;

        #endregion


        public void LoadMailboxes()
        {
            var database = EWSConstants.Config.Database;

            using (EWSDbContext context = new EWSDbContext(database))
            {
                foreach (var sMailbox in context.RoomListRoomEntities.ToList())
                {
                    var addedBox = _mailboxes.AddMailbox(sMailbox.SmtpAddress);
                    if (!addedBox)
                    {
                        _traceListener.Trace("Mailbox Add", $"Failed to add SMTP {sMailbox.SmtpAddress}");
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
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, ewsToken, _traceListener);
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
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, ewsToken, _traceListener);

                            // Add GroupInfo to stack
                            _groups.Add(mailboxInfo.GroupName, groupInfo);
                        }

                        // Add the mailbox to the collection
                        groupInfo.Mailboxes.Add(sMailbox.SmtpAddress);
                    }
                }
            }
        }

        /// <summary>
        /// Enumerate the Groupings, create streaming subscriptions and streaming subscription connections
        /// </summary>
        /// <remarks>You must call LoadMailboxes before this method</remarks>
        public void OpenConnections()
        {
            foreach (var _group in _groups)
            {
                _traceListener.Trace("Open Connections", $"Opening connections for {_group.Key}");

                var groupName = _group.Key;
                var groupInfo = _group.Value;
                if (AddGroupSubscriptions(groupName))
                {
                    _traceListener.Trace("Opened Connections", $"{groupInfo.Mailboxes.Count()} mailboxes primed for StreamingSubscriptions.");
                }
                else
                {
                    _traceListener.Trace("Failed Connections", $"Group {groupInfo.Name} failed in StreamingSubscription events.");
                }
            }

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
                    _traceListener.Trace("Error on Open Connection", $"Error opening streamingsubscriptionconnection for group {_group.Name} MSG {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Adds streamingsubscription to the GroupInfo
        /// </summary>
        /// <param name="Group"></param>
        /// <param name="smtpAddress"></param>
        /// <returns></returns>
        private StreamingSubscription AddSubscription(GroupInfo Group, string smtpAddress)
        {
            if (_subscriptions.Any(email => email.SmtpAddress.Equals(smtpAddress)))
            {
                var email = _subscriptions.FirstOrDefault(s => s.SmtpAddress.Equals(smtpAddress));
                _subscriptions.Remove(email);
            }

            ExchangeService exchange = Group.ExchangeService;
            exchange.Credentials = credentials;
            exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);
            StreamingSubscription subscription = exchange.SubscribeToStreamingNotifications(new FolderId[] { WellKnownFolderName.Calendar },
                    EventType.Created,
                    EventType.Deleted,
                    EventType.Modified,
                    EventType.Moved,
                    EventType.Copied);


            var subscriptionLastMark = default(DateTime?);
            var synchronizationState = string.Empty;
            using (EWSDbContext context = new EWSDbContext(EWSConstants.Config.Database))
            {
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

                var subscriptions = context.SaveChanges();
                _traceListener.Trace("Database", $"Subscriptions {subscriptions} persisted.");
            }

            _subscriptions.Add(new SubscriptionCollection()
            {
                SmtpAddress = smtpAddress,
                Streaming = subscription,
                SubscriptionType = SubscriptionTypeEnum.StreamingSubscription,
                SynchronizationDateTime = subscriptionLastMark,
                SynchronizationState = synchronizationState
            });
            return subscription;
        }

        /// <summary>
        /// Process the GroupInfo and create subscriptions based on AnchorMailbox
        /// </summary>
        /// <param name="groupName">EWS GroupInfo or Dynamic Groupname if Mailboxes > 200</param>
        /// <returns></returns>
        public bool AddGroupSubscriptions(string groupName)
        {
            if (!_groups.ContainsKey(groupName))
                return false;


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
                StreamingSubscription mailboxSubscription = AddSubscription(groupInfo, PrimaryMailbox);
                if (_connections.ContainsKey(groupName))
                {
                    _connections[groupName] = new StreamingSubscriptionConnection(mailboxSubscription.Service, numericUpDownTimeout);
                }
                else
                {
                    _connections.Add(groupName, new StreamingSubscriptionConnection(mailboxSubscription.Service, numericUpDownTimeout));
                }


                SubscribeConnectionEvents(_connections[groupName]);
                _connections[groupName].AddSubscription(mailboxSubscription);
                _traceListener.Trace("Add subscription", String.Format("{0} (primary mailbox) subscription created in group {1}", PrimaryMailbox, groupName));

                // Now add any further subscriptions in this group
                foreach (string sMailbox in groupInfo.Mailboxes.Where(w => !w.Equals(PrimaryMailbox)))
                {
                    try
                    {
                        var localSubscription = AddSubscription(groupInfo, sMailbox);
                        _connections[groupName].AddSubscription(localSubscription);
                        _traceListener.Trace("Add secondary subscription", String.Format("{0} subscription created in group {1}", sMailbox, groupName));
                    }
                    catch (Exception ex)
                    {
                        _traceListener.Trace("Exception", String.Format("ERROR when subscribing {0} in group {1}: {2}", sMailbox, groupName, ex.Message));
                    }

                }
            }
            catch (Exception ex)
            {
                _traceListener.Trace("Exception", String.Format("ERROR when creating subscription connection group {0}: {1}", groupName, ex.Message));
            }
            return true;
        }

        /// <summary>
        /// Reconnect StreamingSubscriptionConnections [if not in a disconnecting or error'd state]
        /// </summary>
        public void Reconnect()
        {
            // Go through our connections and reconnect any that have closed
            _reconnect = false;
            lock (_reconnectLock)  // Prevent this code being run concurrently (i.e. if an event fires in the middle of the processing)
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
                                _traceListener.Trace("Connecting...", String.Format("Re-opened connection for group {0}", connectionGroupName));
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.StartsWith("You must add at least one subscription to this connection before it can be opened"))
                                {
                                    // Try recreating this group
                                    AddGroupSubscriptions(connectionGroupName);
                                }
                                else
                                {
                                    _traceListener.Trace("Exception - reopen", String.Format("Failed to reopen connection: {0}", ex.Message));
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            _traceListener.Trace("Exception", String.Format("Failed to reopen connection: {0}", ex.Message));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Subscribe to events for this connection
        /// </summary>
        /// <param name="connection"></param>
        private void SubscribeConnectionEvents(StreamingSubscriptionConnection connection)
        {
            connection.OnNotificationEvent += connection_OnNotificationEvent;
            connection.OnDisconnect += connection_OnDisconnect;
            connection.OnSubscriptionError += connection_OnSubscriptionError;
        }

        void connection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                _traceListener.Trace("OnSubscriptionError", String.Format("OnSubscriptionError received for {0}: {1}", args.Subscription.Service.ImpersonatedUserId.Id, args.Exception.Message));
            }
            catch
            {
                _traceListener.Trace("Exception", "OnSubscriptionError received");
            }
        }

        void connection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                _traceListener.Trace("connection_OnDisconnect", String.Format("OnDisconnection received for {0}", args.Subscription.Service.ImpersonatedUserId.Id));
            }
            catch
            {
                _traceListener.Trace("Exception", "OnDisconnection received");
            }
            _reconnect = true;  // We can't reconnect in the disconnect event, so we set a flag for the timer to pick this up and check all the connections
        }

        void connection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            foreach (NotificationEvent e in args.Events)
            {
                ProcessNotification(e, args.Subscription);
            }
        }

        /// <summary>
        /// Process events into Service Bus or other messaging queue
        /// </summary>
        /// <param name="e"></param>
        /// <param name="Subscription"></param>
        void ProcessNotification(object e, StreamingSubscription Subscription)
        {
            // We have received a notification

            string sMailbox = Subscription.Service.ImpersonatedUserId.Id;

            if (String.IsNullOrEmpty(sMailbox))
                sMailbox = "Unknown mailbox";
            string sEvent = sMailbox + ": ";

            if (e is ItemEvent)
            {

                sEvent += "Item " + (e as ItemEvent).EventType.ToString() + ": ";
                sEvent += "ItemId = " + (e as ItemEvent).ItemId.UniqueId;
            }
            else if (e is FolderEvent)
            {

                sEvent += "Folder " + (e as FolderEvent).EventType.ToString() + ": ";
                sEvent += "FolderId = " + (e as FolderEvent).FolderId.UniqueId;
            }



        }


        public void DisposeIt()
        {
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

            for (int i = _subscriptions.Count - 1; i >= 0; i--)
            {
                var subscriptionItem = _subscriptions[i];
                var subscription = subscriptionItem.Streaming;
                try
                {
                    subscription.Unsubscribe();
                    _traceListener.Trace("Unsubscribing", String.Format("Unsubscribed from {0}", subscriptionItem.SmtpAddress));
                }
                catch (Exception ex)
                {
                    _traceListener.Trace("Unsubscribing", String.Format("Error when unsubscribing {0}: {1}", subscriptionItem.SmtpAddress, ex.Message));
                }

                _subscriptions.Remove(subscriptionItem);
            }


            _reconnect = false;
        }

        public void Dispose()
        {
            DisposeIt();
        }
    }
}
