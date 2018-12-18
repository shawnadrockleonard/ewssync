using EWS.Common.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace EWSServiceBusSendToO365
{
    /// <summary>
    /// Will process events into O365; could be service bus or a simple ticker
    ///     2 Async operaters in this program
    ///         1. Send Messages to Service bus with Ticker [5 minute intervals] checks local db and writes message
    ///         2. Receive messages from Service bus; operates on threads and messaging in queue
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
            Trace.WriteLine("Capturing Send to O365 ...");


            var p = new Program();

            _handler += new EventHandler(p.ConsoleCtrlCheck);
            SetConsoleCtrlHandler(_handler, true);

            try
            {
                var tasks = new List<System.Threading.Tasks.Task>();

                Trace.WriteLine("In Thread RunAsync....");
                IsDisposed = false;

                // Initialize the Manager with a Token for cancellation
                Messenger = new MessageManager(CancellationTokenSource);

                // Service Bus Connection
                var queueConnection = EWSConstants.Config.ServiceBus.SendToO365;

                // send and tick
                var sendTask = Messenger.SendQueueDatabaseChangesAsync(queueConnection);
                tasks.Add(sendTask);

                // receive queue messages
                var receiveTask = Messenger.ReceiveQueueDatabaseChangesAsync(queueConnection);
                tasks.Add(receiveTask);

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

            // dispose all threads
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
