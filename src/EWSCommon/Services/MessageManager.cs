using EWS.Common.Models;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Azure.ServiceBus;
using Microsoft.Azure.ServiceBus.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Data.Entity;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using EWS.Common.Database;
using Newtonsoft.Json;

namespace EWS.Common.Services
{
    public class MessageManager : IDisposable
    {
        private CancellationTokenSource _cancel { get; set; }

        private AuthenticationResult Ewstoken { get; set; }

        private string MailboxOwner { get; set; }

        private EWSDbContext EwsDatabase { get; set; }

        private bool IsDisposed { get; set; }

        /// <summary>
        /// Database [emulates local app]
        /// </summary>
        private const string SBMessageSyncDb = "SyncDb";
        private const string SBQueueSyncDb = "too365";
        /// <summary>
        /// Streaming Subscription
        /// </summary>
        private const string SBMessageSubscriptionO365 = "SubO365";
        private const string SBQueueSubscriptionO365 = "o365subscription";
        /// <summary>
        /// SyncFolders
        /// </summary>
        private const string SBMessageSyncO365 = "SyncO365";
        private const string SBQueueSyncO365 = "o365sync";

        private Random RandomSeed { get; set; }



        public MessageManager() : this(new CancellationTokenSource(), null)
        {
        }

        public MessageManager(CancellationTokenSource token, AuthenticationResult authenticationResult)
        {
            _cancel = token;
            Ewstoken = authenticationResult;
            MailboxOwner = EWSConstants.Config.Exchange.ImpersonationAcct;
            EwsDatabase = new EWSDbContext(EWSConstants.Config.Database);
            RandomSeed = new Random();
        }

        public void IssueCancellation(CancellationTokenSource cancellationTokenSource)
        {
            _cancel = cancellationTokenSource;
        }

        /// <summary>
        /// Poll localDB events and push service bus events to the Queue
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task SendQueueDatabaseChangesAsync(string queueConnection)
        {
            var sender = new MessageSender(queueConnection, SBQueueSyncDb);

            var EwsService = new EWService(Ewstoken);
            EwsService.SetImpersonation(ConnectingIdType.SmtpAddress, MailboxOwner);

            // Poll the rooms to store locally
            var roomlisting = EwsService.GetRoomListing();

            foreach (var roomlist in roomlisting)
            {
                foreach (var room in roomlist.Value)
                {
                    EntityRoomListRoom databaseRoom = null;
                    if (EwsDatabase.RoomListRoomEntities.Any(s => s.SmtpAddress == room.Address))
                    {
                        databaseRoom = await EwsDatabase.RoomListRoomEntities.FirstOrDefaultAsync(f => f.SmtpAddress == room.Address);
                    }
                    else
                    {
                        databaseRoom = new EntityRoomListRoom()
                        {
                            SmtpAddress = room.Address,
                            Identity = room.Address,
                            RoomList = roomlist.Key
                        };
                        EwsDatabase.RoomListRoomEntities.Add(databaseRoom);
                    }
                }
            }


            var roomchanges = await EwsDatabase.SaveChangesAsync();
            Trace.WriteLine($"Rooms {roomchanges} saved to database.");


            var waitTimer = new TimeSpan(0, 5, 0);
            while (!_cancel.IsCancellationRequested)
            {
                var milliseconds = (int)waitTimer.TotalMilliseconds;

                // whatever you want to happen every 5 minutes
                Trace.WriteLine($"PullRoomReservationChangesAsync({MailboxOwner}) starting at {DateTime.UtcNow.ToShortTimeString()}");


                var i = 0;
                var bookings = EwsDatabase.AppointmentEntities.Where(w => !w.ExistsInExchange || !w.SyncedWithExchange || (w.DeletedLocally && !w.SyncedWithExchange));
                foreach (var booking in bookings)
                {
                    Microsoft.Exchange.WebServices.Data.EventType eventType = Microsoft.Exchange.WebServices.Data.EventType.Deleted;
                    if (!booking.ExistsInExchange)
                    {
                        eventType = Microsoft.Exchange.WebServices.Data.EventType.Created;
                    }
                    else if (!booking.SyncedWithExchange)
                    {
                        eventType = Microsoft.Exchange.WebServices.Data.EventType.Modified;
                    }

                    var ewsbooking = new EWS.Common.Models.UpdatedBooking()
                    {
                        DatabaseId = booking.Id,
                        MailBoxOwnerEmail = booking.OrganizerSmtpAddress,
                        SiteMailBox = booking.Room.SmtpAddress,
                        Subject = booking.Subject,
                        Location = booking.Location,
                        StartUTC = booking.StartUTC,
                        EndUTC = booking.EndUTC,
                        ExchangeId = booking.BookingId,
                        ExchangeChangeKey = booking.BookingChangeKey,
                        BookingReference = booking.BookingReference,
                        ExchangeEvent = eventType
                    };

                    var message = new Message(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(ewsbooking)))
                    {
                        ContentType = "application/json",
                        Label = SBMessageSyncDb,
                        MessageId = i.ToString(),
                        TimeToLive = TimeSpan.FromMinutes(2)
                    };

                    await sender.SendAsync(message);
                    Trace.WriteLine($"{++i}. Sent: Id = {message.MessageId} w/ Subject:{booking.Subject}");
                }

                Trace.WriteLine($"Sent {i} messages");

                Trace.WriteLine($"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                System.Threading.Thread.Sleep(milliseconds);
            }

        }

        /// <summary>
        /// Read Queue for Database Changes [preferably on a separate thread] and Store in Office 365
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task ReceiveQueueDatabaseChangesAsync(string queueConnection)
        {
            var EwsService = new EWService(Ewstoken);

            // Poll the rooms to store locally
            var roomlisting = await EwsDatabase.RoomListRoomEntities.ToListAsync();

            var propertyIds = new List<PropertyDefinitionBase>()
            {
                ItemSchema.Subject,
                AppointmentSchema.Location,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                EWSConstants.RefIdPropertyDef,
                EWSConstants.DatabaseIdPropertyDef
            };

            var receiver = new MessageReceiver(queueConnection, SBQueueSyncDb, ReceiveMode.PeekLock);

            _cancel.Token.Register(() => receiver.CloseAsync());

            // With the receiver set up, we then enter into a simple receive loop that terminates 
            // when the cancellation token if triggered.
            while (!_cancel.Token.IsCancellationRequested)
            {
                try
                {
                    // ask for the next message "forever" or until the cancellation token is triggered
                    var message = await receiver.ReceiveAsync();
                    if (message != null)
                    {
                        if (message.Label != null &&
                           message.ContentType != null &&
                           message.Label.Equals(SBMessageSyncDb, StringComparison.InvariantCultureIgnoreCase) &&
                           message.ContentType.Equals("application/json", StringComparison.InvariantCultureIgnoreCase))
                        {
                            // service bus
                            // #TODO: Read bus events from database and write to O365
                            var booking = JsonConvert.DeserializeObject<EWS.Common.Models.UpdatedBooking>(Encoding.UTF8.GetString(message.Body));
                            Trace.WriteLine($"Msg received: {booking.SiteMailBox} - {booking.Subject}. Cancel status: {booking.ExchangeEvent.ToString("f")}");

                            var eventType = booking.CancelStatus;
                            var eventStatus = booking.ExchangeEvent;
                            var itemId = booking.ExchangeId;
                            var changeKey = booking.ExchangeChangeKey;

                            var appointmentMailbox = booking.SiteMailBox;

                            try
                            {
                                Appointment meeting = null;

                                if (!string.IsNullOrEmpty(itemId))
                                {
                                    var exchangeId = new ItemId(itemId);


                                    if (eventStatus == EventType.Deleted)
                                    {

                                    }
                                    else
                                    {

                                        var apptMeeting = EwsService.GetAppointment(ConnectingIdType.SmtpAddress, booking.MailBoxOwnerEmail, exchangeId, propertyIds);
                                        meeting = apptMeeting.Item;
                                        meeting.Subject = booking.Subject;
                                        meeting.Start = booking.StartUTC;
                                        meeting.End = booking.EndUTC;
                                        meeting.Location = booking.Location;
                                        //meeting.ReminderDueBy = DateTime.Now;
                                        meeting.Save(SendInvitationsMode.SendOnlyToAll);

                                        // Verify that the appointment was created by using the appointment's item ID.
                                        var item = Item.Bind(EwsService.Current, meeting.Id, new PropertySet(propertyIds));
                                        Trace.WriteLine($"Appointment modified: {item.Subject} && ReserveRoomAsync({booking.Location}) completed");

                                        if (item is Appointment)
                                        {
                                            Trace.WriteLine($"Item is Appointment");
                                        }
                                    }
                                }
                                else
                                {
                                    EwsService.SetImpersonation(ConnectingIdType.SmtpAddress, booking.MailBoxOwnerEmail);

                                    meeting = new Appointment(EwsService.Current);
                                    meeting.Resources.Add(booking.SiteMailBox);
                                    meeting.Subject = booking.Subject;
                                    meeting.Start = booking.StartUTC;
                                    meeting.End = booking.EndUTC;
                                    meeting.Location = booking.Location;
                                    meeting.SetExtendedProperty(EWSConstants.RefIdPropertyDef, booking.BookingReference);
                                    meeting.SetExtendedProperty(EWSConstants.DatabaseIdPropertyDef, booking.DatabaseId);
                                    //meeting.ReminderDueBy = DateTime.Now;
                                    meeting.Save(SendInvitationsMode.SendOnlyToAll);

                                    // Verify that the appointment was created by using the appointment's item ID.
                                    var item = Item.Bind(EwsService.Current, meeting.Id, new PropertySet(propertyIds));
                                    Trace.WriteLine($"Appointment created: {item.Subject} && ReserveRoomAsync({booking.Location}) completed");

                                    if (item is Appointment)
                                    {
                                        Trace.WriteLine($"Item is Appointment");
                                    }
                                }

                                // this should exists as this particular message was sent by the database
                                var dbAppointment = EwsDatabase.AppointmentEntities.FirstOrDefault(a => a.Id == booking.DatabaseId);
                                if (dbAppointment != null)
                                {
                                    dbAppointment.ExistsInExchange = true;
                                    dbAppointment.SyncedWithExchange = true;
                                    dbAppointment.ModifiedDate = DateTime.UtcNow;

                                    if (meeting != null)
                                    {
                                        dbAppointment.BookingId = meeting.Id.UniqueId;
                                        dbAppointment.BookingChangeKey = meeting.Id.ChangeKey;
                                    }
                                    else
                                    {
                                        dbAppointment.IsDeleted = true;
                                    }
                                }

                                var appointmentsSaved = EwsDatabase.SaveChanges();
                                Trace.WriteLine($"Saved {appointmentsSaved} rows");
                            }
                            catch (Exception dbex)
                            {
                                Trace.WriteLine($"Error occurred in Appointment creation or Database change {dbex}");
                            }
                            finally
                            {
                                await receiver.CompleteAsync(message.SystemProperties.LockToken);
                            }
                        }
                        else
                        {
                            // purge / log it
                            await receiver.DeadLetterAsync(message.SystemProperties.LockToken);//, "ProcessingError", "Don't know what to do with this message");
                        }
                    }
                }
                catch (ServiceBusException e)
                {
                    if (!e.IsTransient)
                    {
                        Trace.WriteLine(e.Message);
                        throw;
                    }
                }
            }
            await receiver.CloseAsync();
        }

        /// <summary>
        /// Wait for Room Events and send it to the O365 Subscription Service Bus Queue
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <param name="roomSmtp"></param>
        /// <param name="roomEvent"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task SendQueueO365ChangesAsync(string queueConnection, string roomSmtp, ItemEvent roomEvent)
        {
            var sender = new MessageSender(queueConnection, SBQueueSubscriptionO365);

            Trace.WriteLine($"SendQueueO365ChangesAsync({roomSmtp}, {roomEvent.EventType.ToString("f")}) status");
            var i = RandomSeed.Next(1, 15679);

            var ewsbooking = new EventBooking()
            {
                SiteMailBox = roomSmtp,
                EventType = roomEvent.EventType,
                TimeStamp = roomEvent.TimeStamp,
                ExchangeId = roomEvent.ItemId.UniqueId,
                ExchangeChangeKey = roomEvent.ItemId.ChangeKey
            };

            if (roomEvent.OldItemId != null)
            {
                ewsbooking.OldExchangeId = roomEvent.OldItemId.UniqueId;
                ewsbooking.OldExchangeChangeKey = roomEvent.OldItemId.ChangeKey;
            }

            var message = new Message(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(ewsbooking)))
            {
                ContentType = "application/json",
                Label = SBMessageSubscriptionO365,
                MessageId = i.ToString(),
                TimeToLive = TimeSpan.FromMinutes(20)
            };

            await sender.SendAsync(message);
            Trace.WriteLine($"SendQueueO365ChangesAsync({roomSmtp}, {roomEvent.EventType.ToString("f")}) status => sent: Id = {message.MessageId}");
        }

        /// <summary>
        /// Read Queue for O365 Subscription changes and write it to the Database [preferably on a separate thread]
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task ReceiveQueueO365ChangesAsync(string queueConnection)
        {
            Trace.WriteLine($"GetToO365() starting");

            var EwsService = new EWService(Ewstoken);

            var receiver = new MessageReceiver(queueConnection, SBQueueSubscriptionO365, ReceiveMode.PeekLock);

            _cancel.Token.Register(() => receiver.CloseAsync());


            var filterPropertyList = new List<PropertyDefinitionBase>()
            {
                    AppointmentSchema.Location,
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.IsMeeting,
                    AppointmentSchema.IsOnlineMeeting,
                    AppointmentSchema.IsAllDayEvent,
                    AppointmentSchema.IsRecurring,
                    AppointmentSchema.IsCancelled,
                    ItemSchema.IsUnmodified,
                    AppointmentSchema.TimeZone,
                    AppointmentSchema.ICalUid,
                    ItemSchema.ParentFolderId,
                    ItemSchema.ConversationId,
                    AppointmentSchema.ICalRecurrenceId,
                    EWSConstants.RefIdPropertyDef,
                    EWSConstants.DatabaseIdPropertyDef
            };

            // With the receiver set up, we then enter into a simple receive loop that terminates 
            // when the cancellation token if triggered.
            while (!_cancel.Token.IsCancellationRequested)
            {
                try
                {
                    // ask for the next message "forever" or until the cancellation token is triggered
                    var message = await receiver.ReceiveAsync();
                    if (message != null)
                    {
                        if (message.Label != null &&
                           message.ContentType != null &&
                           message.Label.Equals(SBMessageSubscriptionO365, StringComparison.InvariantCultureIgnoreCase) &&
                           message.ContentType.Equals("application/json", StringComparison.InvariantCultureIgnoreCase))
                        {
                            // service bus
                            // #TODO: Read bus events from O365 and write to Database
                            var booking = JsonConvert.DeserializeObject<EWS.Common.Models.EventBooking>(Encoding.UTF8.GetString(message.Body));
                            Trace.WriteLine($"Msg received: {booking.SiteMailBox} status: {booking.EventType.ToString("f")}");

                            var eventType = booking.EventType;
                            var itemId = booking.ExchangeId;
                            var changeKey = booking.ExchangeChangeKey;
                            var appointmentMailbox = booking.SiteMailBox;

                            try
                            {
                                Appointment meeting = null;


                                var dbroom = EwsDatabase.RoomListRoomEntities.FirstOrDefault(f => f.SmtpAddress == appointmentMailbox);


                                if (eventType == EventType.Deleted)
                                {
                                    var entity = EwsDatabase.AppointmentEntities.FirstOrDefault(f => f.BookingId == itemId);
                                    entity.IsDeleted = true;
                                    entity.ModifiedDate = DateTime.UtcNow;
                                    entity.SyncedWithExchange = true;
                                    entity.ExistsInExchange = true;
                                }
                                else
                                {

                                    var appointmentTime = EwsService.GetAppointment(ConnectingIdType.SmtpAddress, appointmentMailbox, itemId, filterPropertyList);

                                    var parentAppointment = EwsService.GetParentAppointment(appointmentTime, filterPropertyList);
                                    meeting = parentAppointment.Item;

                                    var mailboxId = parentAppointment.Organizer.Address;
                                    var refId = parentAppointment.ReferenceId;
                                    var meetingKey = parentAppointment.MeetingKey;

                                    // TODO: move this to the ServiceBus Processing
                                    if ((!string.IsNullOrEmpty(refId) || meetingKey.HasValue)
                                        && EwsDatabase.AppointmentEntities.Any(f => f.BookingReference == refId || f.Id == meetingKey))
                                    {
                                        var entity = EwsDatabase.AppointmentEntities.FirstOrDefault(f => f.BookingReference == refId || f.Id == meetingKey);

                                        entity.EndUTC = meeting.End.ToUniversalTime();
                                        entity.StartUTC = meeting.Start.ToUniversalTime();
                                        entity.ExistsInExchange = true;
                                        entity.IsRecurringMeeting = meeting.IsRecurring;
                                        entity.Location = meeting.Location;
                                        entity.OrganizerSmtpAddress = mailboxId;
                                        entity.Subject = meeting.Subject;
                                        entity.RecurrencePattern = (meeting.Recurrence == null) ? string.Empty : meeting.Recurrence.ToString();
                                        entity.BookingReference = refId;
                                        entity.BookingChangeKey = changeKey;
                                        entity.BookingId = itemId;
                                        entity.ModifiedDate = DateTime.UtcNow;
                                        entity.SyncedWithExchange = true;

                                    }
                                    else
                                    {
                                        var entity = new EntityRoomAppointment()
                                        {
                                            BookingReference = refId,
                                            BookingId = itemId,
                                            BookingChangeKey = changeKey,
                                            EndUTC = meeting.End.ToUniversalTime(),
                                            StartUTC = meeting.Start.ToUniversalTime(),
                                            ExistsInExchange = true,
                                            IsRecurringMeeting = meeting.IsRecurring,
                                            Location = meeting.Location,
                                            OrganizerSmtpAddress = mailboxId,
                                            Subject = meeting.Subject,
                                            RecurrencePattern = (meeting.Recurrence == null) ? string.Empty : meeting.Recurrence.ToString(),
                                            Room = dbroom
                                        };
                                        EwsDatabase.AppointmentEntities.Add(entity);
                                    }
                                }

                                var appointmentsSaved = EwsDatabase.SaveChanges();
                                Trace.WriteLine($"Saved {appointmentsSaved} rows");
                            }
                            catch (Exception dbex)
                            {
                                Trace.WriteLine($"Error occurred in Appointment creation or Database change {dbex}");
                            }
                            finally
                            {
                                await receiver.CompleteAsync(message.SystemProperties.LockToken);
                            }
                        }
                        else
                        {
                            // purge / log it
                            await receiver.DeadLetterAsync(message.SystemProperties.LockToken);//, "ProcessingError", "Don't know what to do with this message");
                        }
                    }
                }
                catch (ServiceBusException e)
                {
                    if (!e.IsTransient)
                    {
                        Trace.WriteLine(e.Message);
                        throw;
                    }
                }
            }
            await receiver.CloseAsync();
        }

        /// <summary>
        /// Wait for Room Events from SyncFolders process and send it to the O365 Sync Service Bus Queue
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <param name="roomSmtp"></param>
        /// <param name="roomEvent"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task SendQueueO365SyncFoldersAsync(string queueConnection, string roomSmtp, ItemEvent roomEvent)
        {
            var sender = new MessageSender(queueConnection, SBQueueSyncO365);

            Trace.WriteLine($"SendQueueO365SyncFoldersAsync({roomSmtp}, {roomEvent.EventType.ToString("f")}) status");
            var i = RandomSeed.Next(1, 15679);

            var ewsbooking = new EventBooking()
            {
                SiteMailBox = roomSmtp,
                EventType = roomEvent.EventType,
                TimeStamp = roomEvent.TimeStamp,
                ExchangeId = roomEvent.ItemId.UniqueId,
                ExchangeChangeKey = roomEvent.ItemId.ChangeKey
            };

            if (roomEvent.OldItemId != null)
            {
                ewsbooking.OldExchangeId = roomEvent.OldItemId.UniqueId;
                ewsbooking.OldExchangeChangeKey = roomEvent.OldItemId.ChangeKey;
            }

            var message = new Message(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(ewsbooking)))
            {
                ContentType = "application/json",
                Label = SBMessageSyncO365,
                MessageId = i.ToString(),
                TimeToLive = TimeSpan.FromMinutes(20)
            };

            await sender.SendAsync(message);
            Trace.WriteLine($"SendQueueO365SyncFoldersAsync({roomSmtp}, {roomEvent.EventType.ToString("f")}) status => sent: Id = {message.MessageId}");
        }

        /// <summary>
        /// Read Queue for O365 Sync Folder events and write it to the database [preferably on a separate thread]
        /// </summary>
        /// <param name="queueConnection"></param>
        /// <returns></returns>
        public async System.Threading.Tasks.Task ReceiveQueueO365SyncFoldersAsync(string queueConnection)
        {
            Trace.WriteLine($"ReceiveQueueO365SyncFoldersAsync() starting");

            var EwsService = new EWService(Ewstoken);

            var receiver = new MessageReceiver(queueConnection, SBQueueSyncO365, ReceiveMode.PeekLock);

            _cancel.Token.Register(() => receiver.CloseAsync());


            var filterPropertyList = new List<PropertyDefinitionBase>()
            {
                    AppointmentSchema.Location,
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.IsMeeting,
                    AppointmentSchema.IsOnlineMeeting,
                    AppointmentSchema.IsAllDayEvent,
                    AppointmentSchema.IsRecurring,
                    AppointmentSchema.IsCancelled,
                    ItemSchema.IsUnmodified,
                    AppointmentSchema.TimeZone,
                    AppointmentSchema.ICalUid,
                    ItemSchema.ParentFolderId,
                    ItemSchema.ConversationId,
                    AppointmentSchema.ICalRecurrenceId,
                    EWSConstants.RefIdPropertyDef,
                    EWSConstants.DatabaseIdPropertyDef
            };

            // With the receiver set up, we then enter into a simple receive loop that terminates 
            // when the cancellation token if triggered.
            while (!_cancel.Token.IsCancellationRequested)
            {
                try
                {
                    // ask for the next message "forever" or until the cancellation token is triggered
                    var message = await receiver.ReceiveAsync();
                    if (message != null)
                    {
                        if (message.Label != null &&
                           message.ContentType != null &&
                           message.Label.Equals(SBMessageSyncO365, StringComparison.InvariantCultureIgnoreCase) &&
                           message.ContentType.Equals("application/json", StringComparison.InvariantCultureIgnoreCase))
                        {

                        }
                        else
                        {
                            // purge / log it
                            await receiver.DeadLetterAsync(message.SystemProperties.LockToken);//, "ProcessingError", "Don't know what to do with this message");
                        }
                    }
                }
                catch (ServiceBusException e)
                {
                    if (!e.IsTransient)
                    {
                        Trace.WriteLine(e.Message);
                        throw;
                    }
                }
            }
            await receiver.CloseAsync();
        }


        #region Helper Exchange Methods

        public string RetreiveInfo(object e)
        {
            // Get more info for the given item.  This will run on it's own thread
            // so that the main program can continue as usual (we won't hold anything up)

            NotificationInfo n = (NotificationInfo)e;

            var service = new EWService(Ewstoken, true);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, n.Mailbox);

            service.Current.Url = n.Service.Url;
            service.Current.TraceFlags = TraceFlags.All;

            var ewsMoreInfoService = service.Current;


            string sEvent = "";
            if (n.Event is ItemEvent)
            {
                sEvent = n.Mailbox + ": Item " + (n.Event as ItemEvent).EventType.ToString() + ": " + MoreItemInfo(n.Event as ItemEvent, ewsMoreInfoService);
            }
            else
                sEvent = n.Mailbox + ": Folder " + (n.Event as FolderEvent).EventType.ToString() + ": " + MoreFolderInfo(n.Event as FolderEvent, ewsMoreInfoService);

            return ShowEvent(sEvent);
        }

        private string ShowEvent(string eventDetails)
        {
            try
            {
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Ex {ex.Message} Error");
            }

            return eventDetails;
        }

        private string MoreItemInfo(ItemEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Item subject=" + GetItemInfo(e.ItemId, service);

            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        private string MoreFolderInfo(FolderEvent e, ExchangeService service)
        {
            string sMoreInfo = "";
            if (e.EventType == EventType.Deleted)
            {
                // We cannot get more info for a deleted item by binding to it, so skip item details
            }
            else
                sMoreInfo += "Folder name=" + GetFolderName(e.FolderId, service);
            if (e.ParentFolderId != null)
            {
                if (!String.IsNullOrEmpty(sMoreInfo)) sMoreInfo += ", ";
                sMoreInfo += "Parent Folder Name=" + GetFolderName(e.ParentFolderId, service);
            }
            return sMoreInfo;
        }

        private string GetItemInfo(ItemId itemId, ExchangeService service)
        {
            // Retrieve the subject for a given item
            string sItemInfo = "";
            Item oItem;
            PropertySet oPropertySet;


            oPropertySet = new PropertySet(ItemSchema.Subject);

            try
            {
                oItem = Item.Bind(service, itemId, oPropertySet);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            if (oItem is Appointment)
            {
                sItemInfo += "Appointment subject=" + oItem.Subject;
                // Show attendee information
                Appointment oAppt = oItem as Appointment;
                sItemInfo += ",RequiredAttendees=" + GetAttendees(oAppt.RequiredAttendees);
                sItemInfo += ",OptionalAttendees=" + GetAttendees(oAppt.OptionalAttendees);
            }
            else
                sItemInfo += "Item subject=" + oItem.Subject;


            return sItemInfo;
        }

        private string GetAttendees(AttendeeCollection attendees)
        {
            if (attendees.Count == 0) return "none";

            string sAttendees = "";
            foreach (Attendee attendee in attendees)
            {
                if (!String.IsNullOrEmpty(sAttendees))
                    sAttendees += ", ";
                sAttendees += attendee.Name;
            }

            return sAttendees;
        }

        private string GetFolderName(FolderId folderId, ExchangeService service)
        {
            // Retrieve display name of the given folder
            try
            {
                Folder oFolder = Folder.Bind(service, folderId, new PropertySet(FolderSchema.DisplayName));
                return oFolder.DisplayName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion


        #region Database Methods


        /// <summary>
        /// Save database changes asynchronously
        /// </summary>
        /// <returns></returns>
        public async System.Threading.Tasks.Task<int> SaveDbChangesAsync()
        {
            return await EwsDatabase.SaveChangesAsync(_cancel.Token);
        }

        #endregion


        public void Dispose()
        {
            if (!IsDisposed)
            {
                // issue any remaining saves
                var rowchanges = EwsDatabase.SaveChanges();
                Trace.WriteLine($"Saved miscellaneous rows {rowchanges}");

                // close database connection
                EwsDatabase.Dispose();
            }

            if (_cancel.IsCancellationRequested)
            {
                try
                {
                    // cancel the token and issue any registered events
                    // if already cancelled then it will throw an exception
                    _cancel.Cancel();
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Cancellation of token failed with message {ex.Message}");
                }
            }
        }
    }
}
