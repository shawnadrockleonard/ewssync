using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class EWService
    {
        internal AppSettings Config { get; set; }

        internal AuthenticationResult Tokens { get; set; }

        private ExchangeService exchangeService { get; set; }

        public EWService()
        {
            Config = EWSConstants.Config;
        }

        public EWService(AuthenticationResult token, bool enableTracing = false) : this()
        {
            Tokens = token;

            exchangeService = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Local)
            {
                Url = new Uri($"{EWSConstants.EWSUrl}/EWS/Exchange.asmx"),
                TraceEnabled = enableTracing,
                TraceFlags = TraceFlags.All,
                Credentials = new OAuthCredentials(Tokens.AccessToken)
            };

        }

        public ExchangeService Current
        {
            get
            {
                return exchangeService;
            }
        }

        public string ImpersonatedId { get { return exchangeService.ImpersonatedUserId.Id; } }

        public void SetImpersonation(ConnectingIdType connectingIdType, string emailAddress)
        {
            exchangeService.ImpersonatedUserId = new ImpersonatedUserId(connectingIdType, emailAddress);
        }

        public async System.Threading.Tasks.Task CreateExchangeServiceAsync(bool enableTrace = false)
        {
            var tokens = await EWSConstants.AcquireTokenAsync();
            Tokens = tokens;

            exchangeService = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Local)
            {
                Url = new Uri($"{EWSConstants.EWSUrl}/EWS/Exchange.asmx"),
                TraceEnabled = enableTrace,
                TraceFlags = TraceFlags.All,
                Credentials = new OAuthCredentials(Tokens.AccessToken)
            };
        }

        /// <summary>
        /// Return rooms from the RoomList
        /// </summary>
        /// <param name="roomName"></param>
        /// <returns></returns>
        public Dictionary<string, List<EmailAddress>> GetRoomListing(string roomName = null)
        {
            ServicePointManager.DefaultConnectionLimit = ServicePointManager.DefaultPersistentConnectionLimit;

            var subs = new Dictionary<string, List<EmailAddress>>();

            exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Config.Exchange.ImpersonationAcct);

            // TODO: Is there a better way to enumerate a la PS Get-MailBox | Where {$_.ResourceType -eq "Room"}
            foreach (var list in exchangeService.GetRoomLists())
            {
                var rooms = new List<EmailAddress>();
                foreach (var room in exchangeService.GetRooms(list))
                {
                    rooms.Add(room);
                }

                subs.Add(list.Address, rooms);
            }

            return subs;
        }

        /// <summary>
        /// Creates pull subscriptiosn for the specific rooms
        /// </summary>
        /// <param name="timeout"></param>
        /// <param name="watermark"></param>
        /// <returns></returns>
        public PullSubscription CreatePullSubscription(ConnectingIdType connectingIdType, string roomAddress, int timeout = 30, string watermark = null)
        {
            ServicePointManager.DefaultConnectionLimit = ServicePointManager.DefaultPersistentConnectionLimit;

            try
            {
                SetImpersonation(connectingIdType, roomAddress);

                var sub = exchangeService.SubscribeToPullNotifications(
                    new FolderId[] { WellKnownFolderName.Calendar },
                    timeout,
                    watermark,
                    EventType.Created, EventType.Deleted, EventType.Modified, EventType.Moved, EventType.Copied);

                Trace.WriteLine($"CreatePullSubscription {sub.Id} to room {roomAddress}");
                return sub;
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
            {
                Trace.WriteLine($"Failed to provision subscription {srex.Message}");
                throw new Exception($"Subscription could not be created for {roomAddress} with MSG:{srex.Message}");
            }
        }

        /// <summary>
        /// Creates streaming subscription for the specific room
        /// </summary>
        /// <param name="connectingIdType"></param>
        /// <param name="roomAddress"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public StreamingSubscription CreateStreamingSubscription(ConnectingIdType connectingIdType, string roomAddress, int timeout = 30)
        {
            ServicePointManager.DefaultConnectionLimit = ServicePointManager.DefaultPersistentConnectionLimit;

            try
            {
                //TODO: Is there a more scalable way so we don't need to subscribe to each room individually?
                SetImpersonation(connectingIdType, roomAddress);

                // TODO: How to reconnect after app failure and get all events since failure occured
                var sub = exchangeService.SubscribeToStreamingNotifications(
                    new FolderId[] { WellKnownFolderName.Calendar },
                    EventType.Created, EventType.Deleted, EventType.Modified, EventType.Moved, EventType.Copied);

                Trace.WriteLine($"CreateStreamingSubscription {sub.Id} to room {roomAddress}");
                return sub;
            }
            catch (Microsoft.Exchange.WebServices.Data.ServiceRequestException srex)
            {
                Trace.WriteLine($"Failed to provision subscription {srex.Message}");
                throw new Exception($"Subscription could not be created for {roomAddress} with MSG:{srex.Message}");
            }
        }

    }
}
