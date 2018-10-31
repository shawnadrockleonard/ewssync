using EWS.Common.Models;
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

        private ExchangeService ExchangeService { get; set; }

        private ExtendedPropertyDefinition CleanGlobalObjectId { get; set; }

        public EWService()
        {
            Config = EWSConstants.Config;
            CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
        }

        public EWService(AuthenticationResult token, bool enableTracing = false) : this()
        {
            Tokens = token;

            ExchangeService = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Local)
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
                return ExchangeService;
            }
        }

        public string ImpersonatedId { get { return ExchangeService.ImpersonatedUserId.Id; } }

        public void SetImpersonation(ConnectingIdType connectingIdType, string emailAddress)
        {
            ExchangeService.ImpersonatedUserId = new ImpersonatedUserId(connectingIdType, emailAddress);
        }

        public async System.Threading.Tasks.Task CreateExchangeServiceAsync(bool enableTrace = false)
        {
            var tokens = await EWSConstants.AcquireTokenAsync();
            Tokens = tokens;

            ExchangeService = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Local)
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

            ExchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Config.Exchange.ImpersonationAcct);

            // TODO: Is there a better way to enumerate a la PS Get-MailBox | Where {$_.ResourceType -eq "Room"}
            foreach (var list in ExchangeService.GetRoomLists())
            {
                var rooms = new List<EmailAddress>();
                foreach (var room in ExchangeService.GetRooms(list))
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
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <returns></returns>
        public PullSubscription CreatePullSubscription(ConnectingIdType connectingIdType, string roomAddress, int timeout = 30, string watermark = null)
        {
            ServicePointManager.DefaultConnectionLimit = ServicePointManager.DefaultPersistentConnectionLimit;

            try
            {
                SetImpersonation(connectingIdType, roomAddress);

                var sub = ExchangeService.SubscribeToPullNotifications(
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
        /// 
        /// </summary>
        /// <param name="connectingIdType"></param>
        /// <param name="roomAddress"></param>
        /// <param name="itemId"></param>
        /// <param name="filterPropertySet"></param>
        /// <returns></returns>
        public AppointmentObjectId GetAppointment(ConnectingIdType connectingIdType, string roomAddress, ItemId itemId, IList<PropertyDefinitionBase> filterPropertySet)
        {
            SetImpersonation(connectingIdType, roomAddress);
            var appointmentTime = Appointment.Bind(ExchangeService, itemId, new PropertySet(filterPropertySet));

            PropertySet psPropSet = new PropertySet(BasePropertySet.FirstClassProperties)
            {
                CleanGlobalObjectId
            };
            appointmentTime.Load(psPropSet);
            appointmentTime.TryGetProperty(CleanGlobalObjectId, out object CalIdVal);


            var objectId = new AppointmentObjectId()
            {
                Id = itemId,
                Item = appointmentTime,
                Base64UniqueId = Convert.ToBase64String((Byte[])CalIdVal),
                ICalUid = appointmentTime.ICalUid,
                Organizer = appointmentTime.Organizer
            };


            return objectId;
        }


        public AppointmentObjectId GetParentAppointment(AppointmentObjectId dependentAppointment, IList<PropertyDefinitionBase> filterPropertySet)
        {
            IList<PropertyDefinitionBase> findPropertyCollection = new List<PropertyDefinitionBase>()
            {
                ItemSchema.DateTimeReceived,
                ItemSchema.Subject,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.IsAllDayEvent,
                AppointmentSchema.IsRecurring,
                AppointmentSchema.IsCancelled,
                AppointmentSchema.TimeZone
            };

            var icalId = dependentAppointment.ICalUid;
            var mailboxId = dependentAppointment.Organizer.Address;


            var objectId = new AppointmentObjectId()
            {
                ICalUid = dependentAppointment.ICalUid,
                Organizer = dependentAppointment.Organizer
            };

            try
            {
                // Initialize the calendar folder via Impersonation
                SetImpersonation(ConnectingIdType.SmtpAddress, mailboxId);

                CalendarFolder AtndCalendar = CalendarFolder.Bind(ExchangeService, new FolderId(WellKnownFolderName.Calendar, mailboxId), new PropertySet());
                SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(CleanGlobalObjectId, dependentAppointment.Base64UniqueId);
                ItemView ivItemView = new ItemView(5)
                {
                    PropertySet = new PropertySet(BasePropertySet.IdOnly, findPropertyCollection)
                };
                FindItemsResults<Item> fiResults = AtndCalendar.FindItems(sfSearchFilter, ivItemView);
                if (fiResults.Items.Count > 0)
                {
                    var objectItem = fiResults.Items.FirstOrDefault();
                    var ownerAppointmentTime = (Appointment)Item.Bind(ExchangeService, objectItem.Id, new PropertySet( filterPropertySet));

                    Trace.WriteLine($"The first {fiResults.Items.Count()} appointments on your calendar from {ownerAppointmentTime.Start.ToShortDateString()} to {ownerAppointmentTime.End.ToShortDateString()}");


                    objectId.Item = ownerAppointmentTime;
                    objectId.Id = ownerAppointmentTime.Id;
                    var props = ownerAppointmentTime.ExtendedProperties.Where(p => (p.PropertyDefinition.PropertySet == DefaultExtendedPropertySet.Meeting));
                    if (props.Any())
                    {
                        objectId.ReferenceId = (string)props.First(p => p.PropertyDefinition.Name == EWSConstants.RefIdPropertyName).Value;
                        objectId.MeetingKey = (int)props.First(p => p.PropertyDefinition.Name == EWSConstants.MeetingKeyPropertyName).Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Error retreiving calendar {mailboxId} msg:{ex.Message}");
            }

            return objectId;
        }
    }
}
