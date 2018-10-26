using System;

namespace EWS.Common.Database
{
    public interface ITraceLogger
    {
        void LogError(Exception ex, string format, params object[] args);
        void LogInformation(string format, params object[] args);
        void LogWarning(string format, params object[] args);
    }
}