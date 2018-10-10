using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSResourceSync
{
    // We use this class to indicate that the inner exception has been logged/traced, but it's
    // wrapped and passed on to the caller to indicate failure
    public class LoggedException : Exception
    {
        public LoggedException(Exception ex)
        {
            System.Diagnostics.Trace.TraceError($"Ex:{ex.Message}");
        }
        public LoggedException(string message, Exception ex) : base(message, ex)
        {
            System.Diagnostics.Trace.TraceError($"Ex:{ex.Message} and user msg:{message}");
        }
    }
}