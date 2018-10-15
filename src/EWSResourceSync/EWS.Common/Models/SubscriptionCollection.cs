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
        public PullSubscription Runningsubscription { get; set; }

        public EntitySubscription DatabaseSubscription { get; set; }


        public string SmtpAddress { get; set; }
    }
}
