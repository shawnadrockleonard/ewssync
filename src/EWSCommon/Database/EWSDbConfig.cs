using System.Data.Entity;

namespace EWS.Common.Database
{
    internal class EWSDbConfig : DbConfiguration
    {
        public EWSDbConfig()
        {
            SetProviderServices("System.Data.SqlClient", System.Data.Entity.SqlServer.SqlProviderServices.Instance);
            AddInterceptor(new EWSDbLogInterceptor());
        }
    }
}
