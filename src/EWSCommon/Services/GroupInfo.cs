using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class GroupInfo
    {
        private string _name = "";
        private string _primaryMailbox = "";
        private List<String> _mailboxes;
        //private List<StreamingSubscriptionConnection> _streamingConnection;
        private ExchangeService _exchangeService = null;
        private ITraceListener _traceListener = null;
        private string _ewsUrl = "";
        private StreamingSubscriptionConnection _connection { get; set; }
        private bool _isConnectionOpen { get; set; }
        private bool _reconnect { get; set; }
        private Dictionary<string, StreamingSubscription> _subscriptions { get; set; }
        private AuthenticationResult ewsToken { get; set; }
        private const int numericUpDownTimeout = 30;


        public GroupInfo(string Name, string PrimaryMailbox, string EWSUrl, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _primaryMailbox = PrimaryMailbox;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(PrimaryMailbox);
            _isConnectionOpen = false;

            _subscriptions = new Dictionary<string, StreamingSubscription>();

            var EwsToken = System.Threading.Tasks.Task.Run(async () =>
            {
                return await EWSConstants.AcquireTokenAsync();
            });

            ewsToken = EwsToken.Result;

        }

        public string Name
        {
            get { return _name; }
        }

        public string PrimaryMailbox
        {
            get { return _primaryMailbox; }
            set
            {
                // If the primary mailbox changes, we need to ensure that it is in the mailbox list also
                _primaryMailbox = value;
                if (!_mailboxes.Contains(_primaryMailbox))
                    _mailboxes.Add(_primaryMailbox);
            }
        }

        public ExchangeService ExchangeService
        {
            get
            {
                if (_exchangeService != null)
                    return _exchangeService;

                // Create exchange service for this group
                ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2013);
                exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _primaryMailbox);
                exchange.HttpHeaders.Add("X-AnchorMailbox", _primaryMailbox);
                exchange.HttpHeaders.Add("X-PreferServerAffinity", "true");
                exchange.Url = new Uri(_ewsUrl);
                exchange.Credentials = new OAuthCredentials(ewsToken.AccessToken);
                if (_traceListener != null)
                {
                    exchange.TraceListener = _traceListener;
                    exchange.TraceFlags = TraceFlags.All;
                    exchange.TraceEnabled = true;
                }
                return exchange;
            }
        }

        public List<String> Mailboxes
        {
            get { return _mailboxes; }
        }

        /// <summary>
        /// The maximum number of mailboxes in a group shouldn't exceed 200, which means that this group may consist of several groups. 
        /// </summary>
        public int NumberOfGroups
        {
            get { return ((_mailboxes.Count / 200)) + 1; }
        }

        public List<List<String>> MailboxesGrouped
        {
            get
            {
                // Return a list of lists (the group split into lists of 200)
                List<List<String>> groupedMailboxes = new List<List<String>>();
                for (int i = 0; i < NumberOfGroups; i++)
                {
                    List<String> mailboxes = _mailboxes.GetRange(i * 200, 200);
                    groupedMailboxes.Add(mailboxes);
                }
                return groupedMailboxes;
            }
        }

        public void OpenSubscription()
        {
            if (_isConnectionOpen || _connection.IsOpen)
                return;

            _connection.Open();
            _isConnectionOpen = true;
        }

        public bool AddGroupSubscriptions()
        {

            if (_isConnectionOpen || _connection.IsOpen)
            {
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
                    _isConnectionOpen = false;
                }
                catch { }
            }

            _connection = new StreamingSubscriptionConnection(ExchangeService, numericUpDownTimeout);
            SubscribeConnectionEvents(_connection);


            // Return the subscription, or create a new one if we don't already have one
            var localSubscription = AddSubscription(PrimaryMailbox);
            _traceListener.Trace("Add subscription", String.Format("{0} (primary mailbox) subscription created in group {1}", PrimaryMailbox, Name));

            // Now add any further subscriptions in this group
            foreach (string sMailbox in Mailboxes.Where(w => !w.Equals(PrimaryMailbox)))
            {
                try
                {
                    localSubscription = AddSubscription(sMailbox);
                    _connection.AddSubscription(localSubscription);
                    _traceListener.Trace("Add secondary subscription", String.Format("{0} subscription created in group {1}", sMailbox, Name));
                }
                catch (Exception ex)
                {
                    _traceListener.Trace("Exception", String.Format("ERROR when subscribing {0} in group {1}: {2}", sMailbox, Name, ex.Message));
                }

            }

            return true;
        }

        private StreamingSubscription AddSubscription(string Mailbox)
        {
            if (_subscriptions.ContainsKey(Mailbox))
                _subscriptions.Remove(Mailbox);

            ExchangeService exchange = ExchangeService;
            exchange.Credentials = ExchangeService.Credentials;
            exchange.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Mailbox);
            StreamingSubscription subscription = exchange.SubscribeToStreamingNotifications(new FolderId[] { WellKnownFolderName.Calendar },
                    EventType.Created,
                    EventType.Deleted,
                    EventType.Modified,
                    EventType.Moved,
                    EventType.Copied);

            _subscriptions.Add(Mailbox, subscription);
            return subscription;
        }

        private void SubscribeConnectionEvents(StreamingSubscriptionConnection connection)
        {
            // Subscribe to events for this connection

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


        private void CloseConnections()
        {

            if (_connection.IsOpen) _connection.Close();

        }
    }
}
