using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common
{
    public class ClassTraceListener : ITraceListener
    {

        public ClassTraceListener()
        {
            try
            {
                System.Diagnostics.Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
                System.Diagnostics.Trace.AutoFlush = true;
            }
            catch { }
        }

        ~ClassTraceListener()
        {
            try
            {
            }
            catch { }
        }

        public void Trace(string traceType, string traceMessage)
        {
            lock (this)
            {
                try
                {
                    System.Diagnostics.Trace.WriteLine(traceMessage);
                    System.Diagnostics.Trace.Flush();
                }
                catch { }
            }
        }
    }
}
