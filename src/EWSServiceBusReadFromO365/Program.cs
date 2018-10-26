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
using System.IO;
using System.Linq;
using System.Data.Entity;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace EWSServiceBusReadFromO365
{
    /// <summary>
    /// Reader from O365 Events from the StreamingSubscription
    /// </summary>
    class Program
    {
        static private System.Threading.CancellationTokenSource CancellationTokenSource = new System.Threading.CancellationTokenSource();
        static private bool IsDisposed { get; set; }
        static MessageManager Messenger { get; set; }

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Capturing Reading from O365 ...");

            _handler += new EventHandler(ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            var p = new Program();

            var service = System.Threading.Tasks.Task.Run(async () =>
            {
                Trace.WriteLine("In Thread run await....");
                await p.RunAsync();

            }, CancellationTokenSource.Token);
            service.Wait();


            Trace.WriteLine("Done.   Press any key to terminate.");
            Console.ReadLine();
        }

        static Program()
        {

        }

        async private System.Threading.Tasks.Task RunAsync()
        {
            IsDisposed = false;

            var Ewstoken = await EWSConstants.AcquireTokenAsync();


            Messenger = new MessageManager(CancellationTokenSource, Ewstoken);



            var queueSubscription = EWSConstants.Config.ServiceBus.O365Subscription;
            var receiveO365Subscriptions = Messenger.ReceiveQueueO365ChangesAsync(queueSubscription);

            var queueSync = EWSConstants.Config.ServiceBus.O365Sync;
            var receiveO365Sync = Messenger.ReceiveQueueO365SyncFoldersAsync(queueSync);


            await System.Threading.Tasks.Task.WhenAll(
                // receive syncfolder events
                receiveO365Sync,
                // receive queue messages
                receiveO365Subscriptions
            );
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
