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
        public int MeetingKey { get; set; }

        /// <summary>
        /// to tell whether booking is in waiting or bboked it can have value like 0,1, 2,3, you can ignore this for POC; ==1 => delete from O365; ==2 - ask for acceptance; ==3 - remove from O365
        /// </summary>
        [DataMember]
        public int CancelStatus { get; set; }

        /// <summary>
        /// this is smtp address of the organiser of appointment
        /// </summary>
        [DataMember]
        public string MailBoxOwnerEmail { get; set; }

        /// <summary>
        /// endpoint of our system.
        /// </summary>
        [DataMember]
        public string ApiURL { get; set; }

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
        public string StartUTC { get; set; }

        /// <summary>
        /// end time of appointment in utc
        /// </summary>
        [DataMember]
        public string EndUTC { get; set; }

        /// <summary>
        /// custom property saved in appointment, Same is  also saved in database which allows us to link appointment with our systen
        /// </summary>
        [DataMember]
        public string BookingRef { get; set; }

        /// <summary>
        /// smtp address of resource mail box
        /// </summary>
        [DataMember]
        public string SiteMailBox { get; set; }
    }
}
