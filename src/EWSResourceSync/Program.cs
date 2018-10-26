using EWS.Common;
using EWS.Common.Models;
using EWS.Common.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;

namespace EWSResourceSync
{
    class Program
    {
        static private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        static private bool IsDisposed { get; set; }
        static MessageManager Messenger { get; set; }
        static AuthenticationResult EwsToken { get; set; }

        private static List<StreamingSubscriptionConnection> _connections = new List<StreamingSubscriptionConnection>();
        private static List<SubscriptionCollection> subscriptions { get; set; }


        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Starting...");

            _handler += new EventHandler(ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            var p = new Program();

            var service = System.Threading.Tasks.Task.Run(async () =>
            {
                Trace.WriteLine("In Thread run await....");
                await p.RunAsync();

            }, CancellationTokenSource.Token);
            service.Wait();

            //hold the console so it doesn’t run off the end
            Console.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }

        static Program()
        {
        }

        async private System.Threading.Tasks.Task RunAsync()
        {
            IsDisposed = false;

            EwsToken = await EWSConstants.AcquireTokenAsync();

            Messenger = new MessageManager(CancellationTokenSource, EwsToken);

            var queueSubscription = EWSConstants.Config.ServiceBus.O365Subscription;
            var impersonationId = EWSConstants.Config.Exchange.ImpersonationAcct;

            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                if (EWSConstants.Config.Exchange.PullEnabled)
                {
                    tasks.Add(PullSubscriptionChangesAsync(queueSubscription, impersonationId));
                }

                tasks.Add(StreamingSubscriptionChangesAsync(queueSubscription, impersonationId));

                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private async System.Threading.Tasks.Task StreamingSubscriptionChangesAsync(string queueConnection, string mailboxOwner)
        {
            Trace.WriteLine($"StreamingSubscriptionChangesAsync({mailboxOwner}) starting");

            var service = new EWService(EwsToken);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            subscriptions = Messenger.CreateStreamingSubscriptionGrouping();

            // Create a streaming connection to the service object, over which events are returned to the client.
            // Keep the streaming connection open for 30 minutes.
            var connection = new StreamingSubscriptionConnection(service.Current, subscriptions.Select(s => s.Streaming), 30);
            var semaphore = new System.Threading.SemaphoreSlim(1);

            connection.OnNotificationEvent += (object sender, NotificationEventArgs args) =>
            {
                string fromEmailAddress = args.Subscription.Service.ImpersonatedUserId.Id;
                Trace.WriteLine($"StreamingSubscriptionChangesAsync received {args.Events.Count()} notification(s)");


                var ewsService = args.Subscription.Service;
                var subscription = subscriptions.FirstOrDefault(k => k.Streaming.Id == args.Subscription.Id);
                var watermark = subscription.Streaming.Watermark;
                var databaseItem = subscription.DatabaseSubscription;
                var email = subscription.SmtpAddress;

                // Process all Events
                // New appt: ev.Type == Created
                // Del apt:  ev.Type == Moved, IsCancelled == true
                // Move apt: ev.Type == Modified, IsUnmodified == false, 
                var evGroups = args.Events.Where(ev => ev is ItemEvent).Select(ev => ((ItemEvent)ev)).OrderBy(x => x.TimeStamp);
                foreach (ItemEvent ev in evGroups)
                {
                    var itemId = ev.ItemId;
                    try
                    {
                        var task = System.Threading.Tasks.Task.Run(async () =>
                        {
                            databaseItem.Watermark = watermark;
                            databaseItem.LastRunTime = DateTime.UtcNow;

                            // Send an item event you can bind to
                            await Messenger.SendQueueO365ChangesAsync(queueConnection, email, ev);

                            await Messenger.SaveDbChangesAsync();
                        });

                        task.Wait(CancellationTokenSource.Token);
                    }
                    catch (ServiceResponseException ex)
                    {
                        Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                        continue;
                    }
                }

            };
            connection.OnDisconnect += (object sender, SubscriptionErrorEventArgs args) =>
            {
                if (args.Exception == null)
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection disconnected");
                else
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection disconnected with exception: {args.Exception.Message}");
                if (CancellationTokenSource.IsCancellationRequested)
                    semaphore.Release();
                else
                {
                    connection.Open();
                    Trace.WriteLine($"ListenToRoomReservationChangesAsync.StreamingSubscriptionConnection re-connected");
                }
            };
            connection.OnSubscriptionError += (object sender, SubscriptionErrorEventArgs args) =>
            {
                if (args.Exception is ServiceResponseException)
                {
                    var exception = args.Exception as ServiceResponseException;
                }
                else if (args.Exception != null)
                {
                    Trace.WriteLine($"OnSubscriptionError() : {args.Exception.Message} Stack Trace : {args.Exception.StackTrace} Inner Exception : {args.Exception.InnerException}");
                }

                connection = (StreamingSubscriptionConnection)sender;
                if (args.Subscription != null)
                {
                    try
                    {
                        connection.RemoveSubscription(args.Subscription);
                    }
                    catch (Exception rex)
                    {
                        Trace.WriteLine($"RemoveSubscriptionException to provision subscription {rex.Message}");
                    }
                }
            };
            semaphore.Wait();
            connection.Open();
            await semaphore.WaitAsync(CancellationTokenSource.Token);

            Trace.WriteLine($"StreamingSubscriptionChangesAsync({mailboxOwner}) exiting");
        }

        private async System.Threading.Tasks.Task PullSubscriptionChangesAsync(string queueConnection, string mailboxOwner)
        {
            Trace.WriteLine($"PullSubscriptionChangesAsync({mailboxOwner}) starting");

            var service = new EWService(EwsToken);
            service.SetImpersonation(ConnectingIdType.SmtpAddress, mailboxOwner);

            subscriptions = Messenger.CreatePullSubscription();
            try
            {
                var waitTimer = new TimeSpan(0, 5, 0);
                while (!CancellationTokenSource.IsCancellationRequested)
                {
                    var milliseconds = (int)waitTimer.TotalMilliseconds;

                    foreach (var item in subscriptions)
                    {
                        bool? ismore = default(bool);
                        do
                        {
                            PullSubscription subscription = item.Pulling;
                            var events = subscription.GetEvents();
                            var watermark = subscription.Watermark;
                            ismore = subscription.MoreEventsAvailable;
                            var databaseItem = item.DatabaseSubscription;
                            var email = item.SmtpAddress;


                            // pull last event from stack TODO: need heuristic for how meetings can be stored
                            var filteredEvents = events.ItemEvents.OrderBy(x => x.TimeStamp);
                            foreach (ItemEvent ev in filteredEvents)
                            {
                                var itemId = ev.ItemId;
                                try
                                {
                                    databaseItem.Watermark = watermark;
                                    databaseItem.LastRunTime = DateTime.UtcNow;

                                    // Send an item event you can bind to
                                    await Messenger.SendQueueO365ChangesAsync(queueConnection, email, ev);
                                    // Save Database changes
                                    await Messenger.SaveDbChangesAsync();
                                }
                                catch (ServiceResponseException ex)
                                {
                                    Trace.WriteLine($"ServiceException: {ex.Message}", "Warning");
                                    continue;
                                }
                            }

                        }
                        while (ismore == true);
                    }

                    Trace.WriteLine($"Sleeping at {DateTime.UtcNow} for {milliseconds} milliseconds...");
                    System.Threading.Thread.Sleep(milliseconds);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Trace.WriteLine($"PullSubscriptionChangesAsync({mailboxOwner}) exiting");
        }




        #region Trap application termination

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);

        private delegate bool EventHandler(CtrlType sig);
        static EventHandler _handler;

        enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        private static bool ConsoleCtrlCheck(CtrlType sig)
        {
            Trace.WriteLine("Exiting system due to external CTRL-C, or process kill, or shutdown");

            // enumerate open subscriptions and close them
            subscriptions.ForEach(subscription =>
            {
                if (subscription.SubscriptionType == EWS.Common.Database.SubscriptionTypeEnum.StreamingSubscription)
                {
                    var f = subscription.Streaming;
                    Trace.WriteLine($"Now unsubscribing for {f.Id}");
                    f.Unsubscribe();
                }
                else if (subscription.SubscriptionType == EWS.Common.Database.SubscriptionTypeEnum.PullSubscription)
                {
                    var f = subscription.Pulling;
                    Trace.WriteLine($"Now unsubscribing for {f.Id}");
                    f.Unsubscribe();
                }

                subscription.DatabaseSubscription.Terminated = true;
                subscription.DatabaseSubscription.LastRunTime = DateTime.UtcNow;
            });

            // should cancel all registered events
            CancellationTokenSource.Cancel();

            // issue into messenger
            Messenger.IssueCancellation(CancellationTokenSource);

            // dispose of the messenger
            if (!IsDisposed)
            {
                // should close out database and issue cancellation to token
                Messenger.Dispose();
            }


            Trace.WriteLine("Cleanup complete");

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }
        #endregion
    }
}
