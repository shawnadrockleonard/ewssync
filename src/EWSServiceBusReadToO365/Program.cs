using EWS.Common;
using EWS.Common.Database;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.ServiceBus.Messaging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EWSServiceBusO365Sync
{
    /// <summary>
    /// Listener to process all of the Service Bus events related to SyncFolder events
    /// </summary>
    class Program
    {
        private static System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();

        static void Main(string[] args)
        {
            Console.WriteLine("Get msgs from O365 SyncFolders");

            Console.WriteLine();
            Console.WriteLine("Listening...");

            var p = new Program();

            var service = System.Threading.Tasks.Task.Run(async () =>
            {
                Console.WriteLine("In Thread run await....");
                await p.RunAsync();

            }, CancellationTokenSource.Token);
            service.Wait();



            Console.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }

        static Program()
        {

        }


        private EWSDbContext EwsDatabase { get; set; }
        private bool IsDisposed { get; set; }
        private EWService EwsService { get; set; }
        private AuthenticationResult Ewstoken { get; set; }


        async private System.Threading.Tasks.Task RunAsync()
        {
            Ewstoken = await EWSConstants.AcquireTokenAsync();

            EwsService = new EWService(Ewstoken);

            EwsDatabase = new EWSDbContext(EWSConstants.Config.Database);


            var queueClient = QueueClient.CreateFromConnectionString(EWSConstants.Config.ServiceBus.O365Sync);
            queueClient.OnMessage((msg) =>
            {

                // #TODO: implement read from bus and write from database
                // Represents events originating in database which could be Create/Update/Delete

                var booking = msg.GetBody<EWS.Common.Models.UpdatedBooking>();
                Console.WriteLine($"Msg received: {booking.SiteMailBox} - {booking.Subject}. Cancel status: {booking.ExchangeEvent.ToString("f")}");


                var itemId = booking.ExchangeItem;
                var appointmentMailbox = booking.SiteMailBox;

                var dbroom = EwsDatabase.RoomListRoomEntities.FirstOrDefaultAsync(f => f.SmtpAddress == appointmentMailbox);

                IList<PropertyDefinitionBase> propertyCollection = new List<PropertyDefinitionBase>()
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
                    AppointmentSchema.Recurrence,
                    EWSConstants.RefIdPropertyDef,
                    EWSConstants.DatabaseIdPropertyDef
                };

                if (booking.ExchangeEvent == EventType.Deleted)
                {

                    var entity = EwsDatabase.AppointmentEntities.FirstOrDefault(f => f.BookingId == itemId.UniqueId);
                    entity.IsDeleted = true;
                    entity.ModifiedDate = DateTime.UtcNow;

                }
                else
                {
                    var appointmentTime = EwsService.GetAppointment(ConnectingIdType.SmtpAddress, booking.SiteMailBox, itemId, propertyCollection);


                    var icalId = appointmentTime.ICalUid;
                    var mailboxId = appointmentTime.Organizer.Address;

                    try
                    {
                        var ownerAppointment = EwsService.GetParentAppointment(appointmentTime, propertyCollection);
                        var ownerAppointmentTime = ownerAppointment.Item;

                        var owernApptId = ownerAppointmentTime.Id;
                        var refId = ownerAppointment.ReferenceId;
                        var meetingKey = ownerAppointment.MeetingKey;


                        // TODO: move this to the ServiceBus Processing
                        if ((!string.IsNullOrEmpty(refId) || meetingKey.HasValue)
                            && EwsDatabase.AppointmentEntities.Any(f => f.BookingReference == refId || f.Id == meetingKey))
                        {
                            var entity = EwsDatabase.AppointmentEntities.FirstOrDefault(f => f.BookingReference == refId || f.Id == meetingKey);

                            entity.EndUTC = ownerAppointmentTime.End.ToUniversalTime();
                            entity.StartUTC = ownerAppointmentTime.Start.ToUniversalTime();
                            entity.ExistsInExchange = true;
                            entity.IsRecurringMeeting = ownerAppointmentTime.IsRecurring;
                            entity.Location = ownerAppointmentTime.Location;
                            entity.OrganizerSmtpAddress = mailboxId;
                            entity.Subject = ownerAppointmentTime.Subject;
                            entity.RecurrencePattern = (ownerAppointmentTime.Recurrence == null) ? string.Empty : ownerAppointmentTime.Recurrence.ToString();
                            entity.BookingReference = refId;
                        }
                        else
                        {
                            var entity = new EntityRoomAppointment()
                            {
                                BookingReference = refId,
                                BookingId = itemId.UniqueId,
                                EndUTC = ownerAppointmentTime.End.ToUniversalTime(),
                                StartUTC = ownerAppointmentTime.Start.ToUniversalTime(),
                                ExistsInExchange = true,
                                IsRecurringMeeting = ownerAppointmentTime.IsRecurring,
                                Location = ownerAppointmentTime.Location,
                                OrganizerSmtpAddress = mailboxId,
                                Subject = ownerAppointmentTime.Subject,
                                RecurrencePattern = (ownerAppointmentTime.Recurrence == null) ? string.Empty : ownerAppointmentTime.Recurrence.ToString(),
                                Room = dbroom.Result
                            };
                            EwsDatabase.AppointmentEntities.Add(entity);
                        }


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error retreiving calendar {mailboxId} msg:{ex.Message}");
                    }
                }

                msg.Complete();


            }, new OnMessageOptions() { AutoComplete = true, MaxConcurrentCalls = 5 });
            queueClient.Close();

        }
    }
}