using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    /// <summary>
    /// This class handles the autodiscover and exposes information needed for grouping
    /// </summary>
    public class Mailboxes
    {
        /// <summary>
        /// Stores information for each mailbox, as returned by autodiscover
        /// </summary>
        private Dictionary<string, MailboxInfo> _mailboxes;  
        /// <summary>
        /// Autodiscovery for exchange attributes
        /// </summary>
        private AutodiscoverService _autodiscover;



        public Mailboxes(OAuthCredentials AutodiscoverCredentials, ITraceListener TraceListener = null)
        {
            _mailboxes = new Dictionary<string, MailboxInfo>();
            _autodiscover = new AutodiscoverService(ExchangeVersion.Exchange2013);  // Minimum version we need is 2013
            _autodiscover.RedirectionUrlValidationCallback = RedirectionCallback;

            if (TraceListener != null)
            {
                _autodiscover.TraceListener = TraceListener;
                _autodiscover.TraceFlags = TraceFlags.All;
                _autodiscover.TraceEnabled = true;
            }

            if (!(AutodiscoverCredentials == null))
                _autodiscover.Credentials = AutodiscoverCredentials;
        }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }


        public List<string> AllMailboxes
        {
            get
            {
                return _mailboxes.Keys.ToList<string>();
            }
        }



        public bool AddMailbox(string SMTPAddress)
        {
            // Perform autodiscover for the mailbox and store the information

            if (_mailboxes.ContainsKey(SMTPAddress))
            {
                // We already have autodiscover information for this mailbox, if it is recent enough we don't bother retrieving it again
                if (_mailboxes[SMTPAddress].IsStale)
                {
                    _mailboxes.Remove(SMTPAddress);
                }
                else
                    return true;
            }

            // Retrieve the autodiscover information
            GetUserSettingsResponse userSettings = null;
            try
            {
                userSettings = GetUserSettings(SMTPAddress);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceInformation(String.Format("Failed to autodiscover for {0}: {1}", SMTPAddress, ex.Message));
                return false;
            }

            // Store the autodiscover result, and check that we have what we need for subscriptions
            MailboxInfo info = new MailboxInfo(SMTPAddress, userSettings);
            if (!info.HaveSubscriptionInformation)
            {
                System.Diagnostics.Trace.TraceInformation(String.Format("Autodiscover succeeded, but EWS Url was not returned for {0}", SMTPAddress));
                return false;
            }

            // Add the mailbox to our list, and if it will be part of a new group add that to the group list (with this mailbox as the primary mailbox)
            _mailboxes.Add(info.SMTPAddress, info);
            return true;
        }

        private GetUserSettingsResponse GetUserSettings(string Mailbox)
        {
            // Attempt autodiscover, with maximum of 10 hops
            // As per MSDN: http://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.autodiscover.autodiscoverservice.getusersettings(v=exchg.80).aspx

            Uri url = null;
            GetUserSettingsResponse response = null;


            for (int attempt = 0; attempt < 10; attempt++)
            {
                _autodiscover.Url = url;
                _autodiscover.EnableScpLookup = (attempt < 2);

                response = _autodiscover.GetUserSettings(Mailbox, UserSettingName.InternalEwsUrl, UserSettingName.ExternalEwsUrl, UserSettingName.GroupingInformation);

                if (response.ErrorCode == AutodiscoverErrorCode.RedirectAddress)
                {
                    return GetUserSettings(response.RedirectTarget);
                }
                else if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl)
                {
                    url = new Uri(response.RedirectTarget);
                }
                else
                {
                    return response;
                }
            }

            throw new Exception("No suitable Autodiscover endpoint was found.");
        }

        public MailboxInfo Mailbox(string SMTPAddress)
        {
            if (_mailboxes.ContainsKey(SMTPAddress))
                return _mailboxes[SMTPAddress];
            return null;
        }
    }
}
