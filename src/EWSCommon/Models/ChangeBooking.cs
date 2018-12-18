using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    [DataContract(Name = "Changes", Namespace = "http://xyz.com/updatedbooking")]
    public class ChangeBooking
    {
        /// <summary>
        /// Gets the date and time when the event occurred.
        /// </summary>
        [DataMember]
        public DateTime? TimeStamp { get; set; }

        [DataMember]
        public ChangeType EventType { get; set; }

        [DataMember]
        public string SiteMailBox { get; set; }

        [DataMember]
        public ConnectingIdType ConnectingType { get; set; }

        /// <summary>
        /// custom property saved in appointment, Same is  also saved in database which allows us to link appointment with our systen
        /// </summary>
        [DataMember]
        public string ExchangeId { get; set; }

        /// <summary>
        /// Exchange ItemId change key
        /// </summary>
        [DataMember]
        public string ExchangeChangeKey { get; set; }
    }
}
