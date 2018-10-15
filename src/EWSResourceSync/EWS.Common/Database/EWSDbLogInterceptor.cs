using System;
using System.Data.Common;
using System.Data.Entity.Infrastructure.Interception;
using System.Diagnostics;

namespace EWS.Common.Database
{
    /// <summary>
    ///     An object that implements this interface can be registered with System.Data.Entity.Infrastructure.Interception.DbInterception
    ///     to receive notifications when Entity Framework executes commands.
    /// </summary>
    public class EWSDbLogInterceptor : IDbInterceptor
    {

        public EWSDbLogInterceptor()
        {
            dblogger = new DefaultUsageLogger();
        }

        public EWSDbLogInterceptor(ITraceLogger logger)
        {
            dblogger = logger;
        }

        internal ITraceLogger dblogger { get; private set; }

        public void NonQueryExecuting(DbCommand command, DbCommandInterceptionContext<int> interceptionContext)
        {
            LogIfNonAsync(command, interceptionContext);
        }

        public void NonQueryExecuted(DbCommand command, DbCommandInterceptionContext<int> interceptionContext)
        {
            LogIfError(command, interceptionContext);
        }

        public void ReaderExecuting(DbCommand command, DbCommandInterceptionContext<DbDataReader> interceptionContext)
        {
            LogIfNonAsync(command, interceptionContext);
        }

        public void ReaderExecuted(DbCommand command, DbCommandInterceptionContext<DbDataReader> interceptionContext)
        {
            LogIfError(command, interceptionContext);
        }

        public void ScalarExecuting(DbCommand command, DbCommandInterceptionContext<object> interceptionContext)
        {
            LogIfNonAsync(command, interceptionContext);
        }

        public void ScalarExecuted(DbCommand command, DbCommandInterceptionContext<object> interceptionContext)
        {
            LogIfError(command, interceptionContext);
        }

        /// <summary>
        /// Will write the parameter names and values to the log file
        /// </summary>
        /// <param name="command"></param>
        [Conditional("DEBUG")]
        [DebuggerStepThrough]
        private void LogParametersIfNonAsync(DbCommand command)
        {
            if (command.Parameters != null)
            {
                var idx = command.Parameters.Count;
                for (var i = 0; i < idx; i++)
                {
                    var prm = command.Parameters[i];
                    dblogger.LogInformation("Parameter:{0} with ParameterValue:{1}", prm.ParameterName, prm.Value);
                }
            }
        }

        private void LogIfNonAsync<TResult>(DbCommand command, DbCommandInterceptionContext<TResult> interceptionContext)
        {
            if (!interceptionContext.IsAsync)
            {
                dblogger.LogInformation("Non-async command used: {0}", command.CommandText);
                LogParametersIfNonAsync(command);
            }
        }

        private void LogInnerException(Exception ex, int idx = 0)
        {
            if (ex.InnerException != null)
            {
                LogInnerException(ex.InnerException, idx++);
            }
            dblogger.LogError(ex, "SQL inner {0} exception {1}", idx, ex.Message);
        }

        private void LogIfError<TResult>(DbCommand command, DbCommandInterceptionContext<TResult> interceptionContext)
        {
            if (interceptionContext.Exception != null)
            {
                dblogger.LogError(interceptionContext.Exception, "Command {0} failed", command.CommandText);
                LogParametersIfNonAsync(command);
                LogInnerException(interceptionContext.Exception);
            }
        }
    }
}
