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


            var p = new Program();

            _handler += new EventHandler(p.ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                Trace.WriteLine("In Thread RunAsync....");
                IsDisposed = false;

                Messenger = new MessageManager(CancellationTokenSource);

                var queueSubscription = EWSConstants.Config.ServiceBus.O365Subscription;
                var receiveO365Subscriptions = Messenger.ReceiveQueueO365ChangesAsync(queueSubscription);
                tasks.Add(receiveO365Subscriptions);

                var queueSync = EWSConstants.Config.ServiceBus.O365Sync;
                var receiveO365Sync = Messenger.ReceiveQueueO365SyncFoldersAsync(queueSync);
                tasks.Add(receiveO365Sync);

                // Wait for each thread
                System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Failed in thread wait {ex.Message} with {ex.StackTrace}");
            }
            finally
            {
                p.Dispose();
            }

            Trace.WriteLine("Done.  Now terminating.");
        }

        static Program()
        {

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

        private bool ConsoleCtrlCheck(CtrlType sig)
        {
            Trace.WriteLine("Exiting system due to external CTRL-C, or process kill, or shutdown");

            // Dispose it
            Dispose();

            //shutdown right away so there are no lingering threads
            Environment.Exit(-1);

            return true;
        }

        private void Dispose()
        {
            Trace.WriteLine("Disposing");
            if (IsDisposed)
                return;

            // should cancel all registered events
            CancellationTokenSource.Cancel();

            // issue into messenger
            Messenger.IssueCancellation(CancellationTokenSource);

            // should close out database and issue cancellation to token
            Messenger.Dispose();

            IsDisposed = true;
            Trace.WriteLine("Cleanup complete");
        }

        #endregion
    }
}
