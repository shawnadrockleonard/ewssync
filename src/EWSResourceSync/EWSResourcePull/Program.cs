using EWS.Common;
using EWS.Common.Database;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace EWSResourcePull
{
    class Program
    {
        const int timeout = 30;
        static string mailboxOwner = "";
        MessageManager _queue;
        private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();

        public static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");


            mailboxOwner = EWSConstants.Config.Exchange.ImpersonationAcct;

            var p = new Program();
            p.Run();

            Console.ReadKey();
        }


        static Program()
        {
        }


        private void Run()
        {
            var service = System.Threading.Tasks.Task.Run(async () =>
                {
                    var resservice = await EWSConstants.AcquireTokenAsync();
                    return resservice;
                }, CancellationTokenSource.Token);
            service.Wait();


            var ewstoken = service.Result;
            var ewsservice = new EWService(ewstoken);
            _queue = new MessageManager(CancellationTokenSource.Token);

            using (var ewsDatabase = new EWSDbContext(EWSConstants.Config.Database))
            {


                try
                {
                    var subscriptions = new Dictionary<PullSubscription, string>();
                    var roomlisting = ewsservice.GetRoomListing();
                    foreach (var roomlist in roomlisting)
                    {
                        foreach (var room in roomlist.Value)
                        {
                            var roomservice = new EWService(ewstoken);

                            var roomsubs = ewsDatabase.SubscriptionEntities.Where(w => w.SmtpAddress == room.Address);
                            EntitySubscription dbSubscription = null;
                            string watermark = null;
                            if (roomsubs.Any(rs => rs.SubscriptionType == SubscriptionTypeEnum.PullSubscription && !rs.Terminated))
                            {
                                dbSubscription = roomsubs.FirstOrDefault(fd => fd.SubscriptionType == SubscriptionTypeEnum.PullSubscription && !fd.Terminated);
                                watermark = dbSubscription.Watermark;
                            }
                            else
                            {
                                dbSubscription = new EntitySubscription()
                                {
                                    LastRunTime = DateTime.UtcNow,
                                    SubscriptionType = SubscriptionTypeEnum.PullSubscription,
                                    SmtpAddress = room.Address
                                };
                                ewsDatabase.SubscriptionEntities.Add(dbSubscription);
                            }

                            var subscription = roomservice.CreatePullSubscription(ConnectingIdType.SmtpAddress, room.Address, timeout, watermark);
                            dbSubscription.Watermark = subscription.Watermark;
                            dbSubscription.Id = subscription.Id;
                            subscriptions.Add(subscription, room.Address);
                        }

                        var rowChanged = ewsDatabase.SaveChanges();
                        Trace.WriteLine($"Pull subscription persisted {rowChanged} rows");
                    }

                    var waitTimer = new TimeSpan(0, 5, 0);
                    while (!CancellationTokenSource.IsCancellationRequested)
                    {
                        var milliseconds = (int)waitTimer.TotalMilliseconds;
                        PullRoomReservationChanges_Tick(ewsDatabase, ewsservice, subscriptions);
                        Trace.WriteLine($"Sleeping for {milliseconds} milliseconds...");
                        System.Threading.Thread.Sleep(milliseconds);
                    }

                    subscriptions.Keys.ToList().ForEach(f => { f.Unsubscribe(); });

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


            }
        }


        private void PullRoomReservationChanges_Tick(EWSDbContext ewsDatabase, EWService ewsService, Dictionary<PullSubscription, string> subs)
        {
            // whatever you want to happen every 5 minutes
            Trace.WriteLine($"PullRoomReservationChangesAsync({mailboxOwner}) starting at {DateTime.UtcNow.ToShortTimeString()}");

            var filterPropertySet = new PropertySet(
                AppointmentSchema.Location,
                AppointmentSchema.Subject,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.IsMeeting,
                AppointmentSchema.IsOnlineMeeting,
                AppointmentSchema.IsAllDayEvent,
                AppointmentSchema.IsRecurring,
                AppointmentSchema.IsCancelled,
                AppointmentSchema.IsUnmodified,
                AppointmentSchema.TimeZone,
                AppointmentSchema.ICalUid,
                AppointmentSchema.ParentFolderId,
                AppointmentSchema.ConversationId,
                AppointmentSchema.ICalRecurrenceId,
                EWSConstants.RefIdPropertyDef,
                EWSConstants.MeetingKeyPropertyDef);

            foreach (var item in subs)
            {
                var subscription = item.Key;
                var events = subscription.GetEvents();
                var watermark = subscription.Watermark;
                var ismore = subscription.MoreEventsAvailable;

                foreach (ItemEvent ev in events.ItemEvents)
                {
                    // Find an item event you can bind to
                    int action = 99;
                    var itemId = ev.ItemId;
                    int meetingKey = 0;
                    string refId = string.Empty;
                    var subscriptionitem = subs.FirstOrDefault(k => k.Key.Id == subscription.Id);
                    try
                    {
                        ewsService.SetImpersonation(ConnectingIdType.SmtpAddress, item.Value);
                        var appointmentTime = (Appointment)Item.Bind(ewsService.Current, itemId);



                        if (ev.EventType == EventType.Created)
                        {
                            action = 0; // created
                        }
                        else if (ev.EventType == EventType.Moved && appointmentTime.IsCancelled)
                        {
                            action = 1; // deleted
                        }
                        else if (ev.EventType == EventType.Modified && !appointmentTime.IsUnmodified)
                        {
                            action = 2; // modified
                        }
                        else
                        {
                            continue;
                        }


                        ExtendedPropertyDefinition CleanGlobalObjectId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, 0x23, MapiPropertyType.Binary);
                        PropertySet psPropSet = new PropertySet(BasePropertySet.FirstClassProperties)
                            {
                                CleanGlobalObjectId
                            };
                        appointmentTime.Load(psPropSet);
                        appointmentTime.TryGetProperty(CleanGlobalObjectId, out object CalIdVal);

                        var icalId = appointmentTime.ICalUid;
                        var mailboxId = appointmentTime.Organizer.Address;

                        // Initialize the calendar folder via Impersonation
                        ewsService.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxId);

                        try
                        {
                            CalendarFolder AtndCalendar = CalendarFolder.Bind(ewsService.Current, new FolderId(WellKnownFolderName.Calendar, mailboxId), filterPropertySet);
                            SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(CleanGlobalObjectId, Convert.ToBase64String((Byte[])CalIdVal));
                            ItemView ivItemView = new ItemView(5)
                            {
                                PropertySet = filterPropertySet
                            };
                            FindItemsResults<Item> fiResults = AtndCalendar.FindItems(sfSearchFilter, ivItemView);
                            if (fiResults.Items.Count > 0)
                            {
                                var filterApt = fiResults.Items.FirstOrDefault() as Appointment;
                                Trace.WriteLine($"The first {fiResults.Items.Count()} appointments on your calendar from {filterApt.Start.ToShortDateString()} to {filterApt.End.ToShortDateString()}");

                                var props = filterApt.ExtendedProperties.Where(p => (p.PropertyDefinition.PropertySet == DefaultExtendedPropertySet.Meeting));
                                if (props.Any())
                                {
                                    refId = (string)props.First(p => p.PropertyDefinition.Name == EWSConstants.RefIdPropertyName).Value;
                                    meetingKey = (int)props.First(p => p.PropertyDefinition.Name == EWSConstants.MeetingKeyPropertyName).Value;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine($"Error retreiving calendar {mailboxId} msg:{ex.Message}");
                        }

                        var task = System.Threading.Tasks.Task.Run(async () =>
                        {
                            await _queue.SendFromO365Async(ewsService.ImpersonatedId, appointmentTime, action);
                        });

                        task.Wait(CancellationTokenSource.Token);
                    }
                    catch (ServiceResponseException ex)
                    {
                        Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                        continue;
                    }
                }
            }
        }
    }
}
