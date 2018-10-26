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
    public class UpdatedBooking
    {
        /// <summary>
        /// this is the unique key stored in Database to identify booking in our system, you can ignore this for POC; >0 existing meeting
        /// </summary>
        [DataMember]
        public int? DatabaseId { get; set; }

        /// <summary>
        /// to tell whether booking is in waiting or booked it can have value like 
        /// </summary>
        [DataMember]
        public BookingChangeEnum CancelStatus { get; set; }

        [DataMember]
        public EventType ExchangeEvent { get; set; }

        /// <summary>
        /// this is smtp address of the organiser of appointment
        /// </summary>
        [DataMember]
        public string MailBoxOwnerEmail { get; set; }

        /// <summary>
        /// smtp address of resource mail box
        /// </summary>
        [DataMember]
        public string SiteMailBox { get; set; }

        /// <summary>
        /// subject of appointment
        /// </summary>
        [DataMember]
        public string Subject { get; set; }

        /// <summary>
        /// location of appointment
        /// </summary>
        [DataMember]
        public string Location { get; set; }

        /// <summary>
        /// strat time of appointment in utc
        /// </summary>
        [DataMember]
        public DateTime StartUTC { get; set; }

        /// <summary>
        /// end time of appointment in utc
        /// </summary>
        [DataMember]
        public DateTime EndUTC { get; set; }

        /// <summary>
        /// custom property saved in appointment, Same is  also saved in database which allows us to link appointment with our systen
        /// </summary>
        [DataMember]
        public string BookingReference { get; set; }

        /// <summary>
        /// Exchange ItemId Unique Id
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
