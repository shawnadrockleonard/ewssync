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
        private readonly string _name = "";
        private string _primaryMailbox = "";
        private List<string> _mailboxes;
        //private List<StreamingSubscriptionConnection> _streamingConnection;
        private ExchangeService _exchangeService = null;
        private ITraceListener _traceListener = null;
        private readonly string _ewsUrl = "";
        private bool _isConnectionOpen = false;
        private AuthenticationResult ewsToken { get; set; }


        public GroupInfo(string Name, string PrimaryMailbox, string EWSUrl, AuthenticationResult authentication, ITraceListener TraceListener = null)
        {
            // initialise the group information
            _name = Name;
            _primaryMailbox = PrimaryMailbox;
            _ewsUrl = EWSUrl;
            _traceListener = TraceListener;
            _mailboxes = new List<String>();
            _mailboxes.Add(PrimaryMailbox);

            ewsToken = authentication;
        }

        public string Name
        {
            get { return _name; }
        }

        /// <summary>
        /// Represents a Dynamic Group name [not from Exchange]
        /// </summary>
        public string GroupInfoName { get; set; }

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
                _exchangeService = new ExchangeService(ExchangeVersion.Exchange2013)
                {
                    ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _primaryMailbox)
                };
                _exchangeService.HttpHeaders.Add("X-AnchorMailbox", _primaryMailbox);
                _exchangeService.HttpHeaders.Add("X-PreferServerAffinity", "true");
                _exchangeService.Url = new Uri(_ewsUrl);
                _exchangeService.Credentials = new OAuthCredentials(ewsToken.AccessToken);
                if (_traceListener != null)
                {
                    _exchangeService.TraceListener = _traceListener;
                    _exchangeService.TraceFlags = TraceFlags.All;
                    _exchangeService.TraceEnabled = true;
                }

                return _exchangeService;
            }
        }


        public List<string> Mailboxes
        {
            get { return _mailboxes; }
        }

        /// <summary>
        /// Indicator that Grouping has reached EWS capacity
        /// </summary>
        public bool HasReachedThreshold
        {
            get
            {
                if(_mailboxes == null || !_mailboxes.Any())
                {
                    return false;
                }

                return _mailboxes.Count() > 199;
            }
        }


        public bool IsConnectionOpen
        {
            get { return _isConnectionOpen; }
            set { _isConnectionOpen = value; }
        }

    }
}
