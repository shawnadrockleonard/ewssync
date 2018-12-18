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
        public EntitySubscription()
        {
        }

        [Key()]
        public int Id { get; set; }

        /// <summary>
        /// Email Address to which the subscription pertains
        /// </summary>
        [MaxLength(255)]
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Watermark for individual subscription
        /// </summary>
        [MaxLength(2048)]
        public string Watermark { get; set; }

        /// <summary>
        /// Watermark from previous run
        /// </summary>
        [MaxLength(2048)]
        public string PreviousWatermark { get; set; }

        /// <summary>
        /// Last Subscription event polling
        /// </summary>
        public DateTime LastRunTime { get; set; }
    }
}
