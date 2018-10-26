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
                await SendRoomReservationChanges_Tick(sender);
                Trace.WriteLine($"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                System.Threading.Thread.Sleep(milliseconds);
            }

        }

        internal async System.Threading.Tasks.Task SendRoomReservationChanges_Tick(MessageSender sender)
        {
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
        }


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



        public async System.Threading.Tasks.Task SendQueueO365ChangesAsync(string queueConnection, string roomSMPT, AppointmentObjectId apt, EventType status)
        {
            var sender = new MessageSender(queueConnection, SBQueueSubscriptionO365);

            var originalAppt = apt.Item;
            Trace.WriteLine($"SendFromO365Async({roomSMPT}, {originalAppt.Subject}) starting");
            var i = RandomSeed.Next(1, 15679);

            var ewsbooking = new UpdatedBooking()
            {
                DatabaseId = apt.MeetingKey,
                MailBoxOwnerEmail = apt.Item.Organizer.Address,
                SiteMailBox = roomSMPT,
                Subject = originalAppt.Subject,
                Location = originalAppt.Location,
                StartUTC = originalAppt.Start,
                EndUTC = originalAppt.End,
                ExchangeId = apt.Id.UniqueId,
                ExchangeChangeKey = apt.Id.ChangeKey,
                BookingReference = apt.ReferenceId,
                ExchangeEvent = status
            };

            var message = new Message(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(ewsbooking)))
            {
                ContentType = "application/json",
                Label = SBMessageSubscriptionO365,
                MessageId = i.ToString(),
                TimeToLive = TimeSpan.FromMinutes(2)
            };

            await sender.SendAsync(message);
            Trace.WriteLine($"SendFromO365Async() exiting {status}: {ewsbooking.Subject} sent: Id = {message.MessageId}");
        }



        public async System.Threading.Tasks.Task ReceiveQueueO365ChangesAsync(string queueConnection)
        {
            Trace.WriteLine($"GetToO365() starting");

            var receiver = new MessageReceiver(queueConnection, SBQueueSubscriptionO365, ReceiveMode.PeekLock);

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
                           message.Label.Equals(SBMessageSubscriptionO365, StringComparison.InvariantCultureIgnoreCase) &&
                           message.ContentType.Equals("application/json", StringComparison.InvariantCultureIgnoreCase))
                        {
                            // service bus
                            // #TODO: Read bus events from O365 and write to Database
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
