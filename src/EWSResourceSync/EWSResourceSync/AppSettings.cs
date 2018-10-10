using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EWSResourceSync
{
    /// <summary>
    /// Enables the query/selection of AppSettings from the App.config | web.config
    /// </summary>
    public class AppSettings
    {
        private static readonly string appFile = string.Empty;
        private static readonly AppSettings current;

        static AppSettings()
        {
            var path = new System.IO.DirectoryInfo(Environment.CurrentDirectory);
            appFile = ConfigurationManager.AppSettings["jsonfile"];
            System.Diagnostics.Trace.TraceInformation($"Running in {path} now reading {appFile} json file into memory....");
            current = (new AppSettings()).Get();
        }

        /// <summary>
        /// The current instance.
        /// </summary>
        public static AppSettings Current
        {
            get { return AppSettings.current; }
        }
        
        public AppSettings Get()
        {
            AppSettings item = null;
            try
            {
                var result = System.IO.File.ReadAllText(appFile);

                item = JsonConvert.DeserializeObject<AppSettings>(result);
                item.AuthCert = ReadCertificateFromStore(item.AzureAD.Thumbprint);
                if (item.AuthCert == null)
                    throw new ApplicationException("Certificate not found");
            }
            catch (Exception ex)
            {
                throw new LoggedException(ex);
            }

            return item;
        }

        public X509Certificate2 AuthCert { get; private set; }

        private X509Certificate2 ReadCertificateFromStore(string thumbprint)
        {
            X509Certificate2 cert = null;
            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            X509Certificate2Collection certCollection = store.Certificates;
            X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
            X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindByThumbprint, thumbprint, false);
            cert = signingCert.OfType<X509Certificate2>().OrderByDescending(c => c.NotBefore).FirstOrDefault();
            store.Close();
            return cert;
        }


        [JsonProperty(PropertyName = "aad")]
        public SettingsAzureAD AzureAD { get; set; }

        [JsonProperty(PropertyName = "exo")]
        public SettingsExchange Exchange { get; set; }

        [JsonProperty(PropertyName = "queues")]
        public SettingsQueues ServiceBus { get; set; }
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
    }

    public class SettingsQueues
    {
        public string SendToO365 { get; set; }

        public string ReadToO365 { get; set; }

        public string SendFromO365 { get; set; }

        public string ReadFromO365 { get; set; }
    }
}
