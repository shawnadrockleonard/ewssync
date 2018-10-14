using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common
{
    /// <summary>
    /// Enables the query/selection of AppSettings from the App.config | web.config
    /// </summary>
    public class AppSettings
    {
        [JsonProperty(PropertyName = "aad")]
        public SettingsAzureAD AzureAD { get; set; }

        [JsonProperty(PropertyName = "exo")]
        public SettingsExchange Exchange { get; set; }

        [JsonProperty(PropertyName = "queues")]
        public SettingsQueues ServiceBus { get; set; }

        [JsonProperty(PropertyName = "database")]
        public string Database { get; set; }
    }

    public class SettingsAzureAD
    {
        public string TenantName { get; set; }

        public string AppId { get; set; }

        public string Thumbprint { get; set; }

        public string AuthPoint { get; set; }

        [JsonIgnore()]
        public string Authority
        {
            get
            {
                return $"{AuthPoint}/{TenantName}";
            }
        }
    }

    public class SettingsExchange
    {
        public string ServerName { get; set; }

        public int BatchSize { get; set; }

        public string ImpersonationAcct { get; set; }
    }

    public class SettingsQueues
    {
        public string SendToO365 { get; set; }

        public string ReadToO365 { get; set; }

        public string SendFromO365 { get; set; }

        public string ReadFromO365 { get; set; }
    }
}
