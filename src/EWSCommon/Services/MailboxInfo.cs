using Microsoft.Exchange.WebServices.Autodiscover;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    /// <summary>
    /// This class stores the autodiscover information for a mailbox and exposes the subscription/grouping data
    /// </summary>
    public class MailboxInfo
    {
        /// <summary>
        /// Set when the class is instantiated, so that we know when we obtained the information
        /// </summary>
        private DateTime _timeInfoSet; 
        public string SMTPAddress { get; private set; }
        public string EwsUrl { get; private set; }
        public string GroupingInformation { get; private set; }
        public string Watermark { get; set; }

        public MailboxInfo()
        {
            _timeInfoSet = DateTime.Now;
        }

        public MailboxInfo(string Mailbox, GetUserSettingsResponse UserSettings) : this()
        {
            SMTPAddress = Mailbox;
            try
            {
                EwsUrl = (string)UserSettings.Settings[UserSettingName.ExternalEwsUrl];
            }
            catch
            {
                try
                {
                    EwsUrl = (string)UserSettings.Settings[UserSettingName.InternalEwsUrl];
                }
                catch { }
            }
            try
            {
                GroupingInformation = (string)UserSettings.Settings[UserSettingName.GroupingInformation];
            }
            catch { }

            if (String.IsNullOrEmpty(GroupingInformation))
            {
                GroupingInformation = "all";
            }
        }

        public bool HaveSubscriptionInformation
        {
            get
            {
                return (!String.IsNullOrEmpty(EwsUrl) && !String.IsNullOrEmpty(GroupingInformation));
            }
        }

        public string GroupName
        {
            get
            {
                return String.Format("{0}{1}", EwsUrl, GroupingInformation);
            }
        }

        public bool IsStale
        {
            // We assume that the information is stale if it is over 24 hours old
            get
            {
                return DateTime.Now.Subtract(_timeInfoSet) > new TimeSpan(1, 0, 0, 0);
            }
        }
    }
}
