using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class EWSConstants
    {
        private static readonly string appFile = string.Empty;

        public static readonly string EWSUrl = "";
        public static readonly string EWSAppId = "";
        public static readonly int GetUserAvailabilityBatchSize = 75;

        public const string RefIdPropertyName = "X-AptRefId";
        public static readonly ExtendedPropertyDefinition RefIdPropertyDef;
        public const string MeetingKeyPropertyName = "X-MeetingKey";
        public static readonly ExtendedPropertyDefinition MeetingKeyPropertyDef;


        static EWSConstants()
        {

            RefIdPropertyDef = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, // InternetHeaders,
                                    RefIdPropertyName,
                                    MapiPropertyType.String);
            MeetingKeyPropertyDef = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, // InternetHeaders,
                                    MeetingKeyPropertyName,
                                    MapiPropertyType.Integer);

            var path = new System.IO.DirectoryInfo(Environment.CurrentDirectory);
            appFile = ConfigurationManager.AppSettings["jsonfile"];
            System.Diagnostics.Trace.TraceInformation($"Running in {path} now reading {appFile} json file into memory....");
            GetAppSettingsFromFile();

            EWSUrl = $"https://{current.Exchange.ServerName}";
            EWSAppId = current.AzureAD.AppId;
            GetUserAvailabilityBatchSize = current.Exchange.BatchSize;
        }

        private static AppSettings current;
        /// <summary>
        /// The current instance.
        /// </summary>
        public static AppSettings Config
        {
            get { return current; }
        }

        public static void GetAppSettingsFromFile()
        {
            try
            {
                var result = System.IO.File.ReadAllText(appFile);

                current = JsonConvert.DeserializeObject<AppSettings>(result);
                EWSAppAuthCert = ReadCertificateFromStore(current.AzureAD.Thumbprint);
            }
            catch (Exception ex)
            {
                throw new LoggedException(ex);
            }


            if (EWSAppAuthCert == null)
            {
                throw new ApplicationException("Certificate not found");
            }
        }

        public static X509Certificate2 EWSAppAuthCert { get; private set; }

        private static X509Certificate2 ReadCertificateFromStore(string thumbprint)
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

        /// <summary>
        /// Return Tokens for calling Exchange Web Services
        /// </summary>
        /// <returns></returns>
        public static async Task<AuthenticationResult> AcquireTokenAsync()
        {
            //Trace.WriteLine($"CreateExchangeServiceAsync({impersonatedUser}, {enableTrace})");
            var auth = new AuthenticationContext(Config.AzureAD.Authority);
            var tokens = await auth.AcquireTokenAsync(EWSUrl, new ClientAssertionCertificate(EWSAppId, EWSAppAuthCert));
            return tokens;
        }
    }
}
