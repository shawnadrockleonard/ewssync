using Microsoft.Exchange.WebServices.Data;
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

        public GroupInfo(string Name, string PrimaryMailbox, string EWSUrl, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _primaryMailbox = PrimaryMailbox;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(PrimaryMailbox);
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
    }
}
