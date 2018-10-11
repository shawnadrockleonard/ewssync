using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class EWService
    {
        private static readonly string EWSUrl = "";
        private static readonly string EWSAppId = "";
        private static readonly X509Certificate2 EWSAppAuthCert;
        private static readonly int GetUserAvailabilityBatchSize = 75;
        static AppSettings config;

        public const string RefIdPropertyName = "X-AptRefId";
        public static readonly ExtendedPropertyDefinition RefIdPropertyDef;
        public const string MeetingKeyPropertyName = "X-MeetingKey";
        public static readonly ExtendedPropertyDefinition MeetingKeyPropertyDef;
        public static readonly Guid _myPropertySetId = new Guid("{DAD02742-32A0-406E-950E-4957E5A394E9}");
        public static PropertySet extendedProperties;

        static EWService()
        {
            config = AppSettings.Current;
            EWSUrl = $"https://{config.Exchange.ServerName}";
            GetUserAvailabilityBatchSize = config.Exchange.BatchSize;
            EWSAppId = config.AzureAD.AppId;
            EWSAppAuthCert = config.AuthCert;

            RefIdPropertyDef = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, // InternetHeaders,
                                    RefIdPropertyName,
                                    MapiPropertyType.String);
            MeetingKeyPropertyDef = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Meeting, // InternetHeaders,
                                    MeetingKeyPropertyName,
                                    MapiPropertyType.Integer);

            //RefIdPropertyDef = new ExtendedPropertyDefinition(
            //    _myPropertySetId, "RefId", MapiPropertyType.String);
            //MeetingKeyPropertyDef = new ExtendedPropertyDefinition(
            //    _myPropertySetId, "MeetingKey", MapiPropertyType.Integer);


            extendedProperties = new PropertySet(BasePropertySet.FirstClassProperties)
            {
                RefIdPropertyDef,
                MeetingKeyPropertyDef
            };
        }

        /// <summary>
        /// The current instance.
        /// </summary>
        public static AppSettings Config
        {
            get { return config; }
        }

        public static async Task<ExchangeService> CreateExchangeServiceAsync(bool enableTrace = false)
        {
            //Trace.WriteLine($"CreateExchangeServiceAsync({impersonatedUser}, {enableTrace})");
            var auth = new AuthenticationContext(config.AzureAD.Authority);
            var tokens = await auth.AcquireTokenAsync(EWSUrl, new ClientAssertionCertificate(EWSAppId, EWSAppAuthCert));

            ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2013, TimeZoneInfo.Local);
            exchangeService.Url = new Uri($"{EWSUrl}/EWS/Exchange.asmx");
            exchangeService.TraceEnabled = true;
            exchangeService.TraceFlags = TraceFlags.All;
            exchangeService.Credentials = new OAuthCredentials(tokens.AccessToken);
            exchangeService.TraceEnabled = enableTrace;

            //Trace.WriteLine($"CreateExchangeServiceAsync({impersonatedUser}, {enableTrace} completed");

            return exchangeService;
        }
    }
}
