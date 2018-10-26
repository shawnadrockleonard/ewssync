using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    [DataContract(Name = "Booking", Namespace = "http://xyz.com/updatedbooking")]
    public class BlockBooking
    {
        /// <summary>
        /// this is the unique key stored in Database to identify booking in our system, you can ignore this for POC; >0 existing meeting
        /// </summary>
        [DataMember]
        public int? DatabaseId { get; set; }

        [DataMember]
        public string MailBoxOwnerEmail { get; set; }

        [DataMember]
        public DateTime StartUTC { get; set; }

        [DataMember]
        public DateTime EndUTC { get; set; }

        [DataMember]
        public string Subject { get; set; }

        [DataMember]
        public string Location { get; set; }


        [DataMember]
        public string SiteMailBox { get; set; }

        /// <summary>
        /// recurrence pattern of recuuring appointments
        /// </summary>
        [DataMember]
        public string RecurrencePattern { get; set; } 


        //public List<BookingException> Exceptions { get; set; } // exception for recurring appointments

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
