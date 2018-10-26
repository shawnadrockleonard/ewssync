using EWS.Common.Database;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    public class SubscriptionCollection
    {
        public PullSubscription Pulling { get; set; }

        public StreamingSubscription Streaming { get; set; }

        public EntitySubscription DatabaseSubscription { get; set; }

        public string SmtpAddress { get; set; }

        /// <summary>
        /// Type of subscription
        /// </summary>
        public SubscriptionTypeEnum SubscriptionType { get; set; }
    }
}
