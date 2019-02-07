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
using System.Threading;
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
        private static readonly AutoResetEvent AutoResetEvent = new AutoResetEvent(false);
        static private bool IsDisposed { get; set; }
        static MessageManager Messenger { get; set; }

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.AutoFlush = true;
            Trace.WriteLine("Capturing Reading from O365 ...");

            var p = new Program();

            var _handler = new EventHandler(p.ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);


            Task.Factory.StartNew(() =>
            {
                Trace.WriteLine("In Thread run await....");
                await p.RunAsync();
                p.DisposeWithToken();

            }, CancellationTokenSource.Token);


            Console.CancelKeyPress += (object sender, ConsoleCancelEventArgs e) =>
            {
                p.OnCancelKeyPress(sender, e);
            };

            AutoResetEvent.WaitOne();
            MessageWrite("After WaitOne fired event");
        }

        static Program()
        {

        }

        private static void MessageWrite(string message)
        {
            System.Diagnostics.Trace.TraceInformation($"{message} at {DateTime.UtcNow}");
            Console.WriteLine(message);
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

        private void OnCancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            MessageWrite("CTRL-C event fired.");
            // Dispose it
            DisposeWithToken();
            e.Cancel = true;
            AutoResetEvent.Set();
        }

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);

        private delegate bool EventHandler(CtrlType sig);

        enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        private bool ConsoleCtrlCheck(CtrlType sig)
        {
            MessageWrite("CTRL-C, or process kill, or shutdown");

            // Dispose it
            DisposeWithToken();

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }


        private void DisposeWithToken()
        {
            Trace.WriteLine("Exiting system due to external CTRL-C, or process kill, or shutdown");
            if (IsDisposed)
                return;

            // should cancel all registered events
            if (CancellationTokenSource.IsCancellationRequested)
                CancellationTokenSource.Cancel();

            // issue into messenger
            Messenger.IssueCancellation(CancellationTokenSource);

            // should close out database and issue cancellation to token
            Messenger.Dispose();

            IsDisposed = true;
            MessageWrite("Cleanup complete");
        }

        #endregion
    }
}
