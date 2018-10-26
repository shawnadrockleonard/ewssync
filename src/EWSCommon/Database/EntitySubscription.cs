using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Database
{
    [Table("Subscriptions", Schema = "dbo")]
    public class EntitySubscription
    {
        /// <summary>
        /// Subscription ID
        /// </summary>
        [Key()]
        public string Id { get; set; }

        [MaxLength(2048)]
        public string Watermark { get; set; }

        [MaxLength(255)]
        public string SmtpAddress { get; set; }


        public DateTime LastRunTime { get; set; }

        /// <summary>
        /// Type of subscription
        /// </summary>
        public SubscriptionTypeEnum SubscriptionType { get; set; }

        /// <summary>
        /// Is the subscription closed by thread or cancelled
        /// </summary>
        public bool Terminated { get; set; }
    }
}
