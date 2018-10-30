using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class GroupInfoSubscriptionConnection
    {
        public string GroupName { get; set; }

        private StreamingSubscriptionConnection connection { get; set; }

        public StreamingSubscriptionConnection GetConnection()
        {
            return connection;
        }

        public void SetConnection(StreamingSubscriptionConnection value)
        {
            connection = value;
        }

        public void AddSubscription(StreamingSubscription subscription)
        {
            connection.AddSubscription(subscription);
        }
    }
}
