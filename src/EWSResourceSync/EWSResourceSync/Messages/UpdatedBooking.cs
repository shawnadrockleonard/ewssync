using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EWSResourceSync
{
    [DataContract(Name = "Booking", Namespace = "http://trimble.com/updatedbooking")]
    public class UpdatedBooking
    {
        [DataMember]
        public int MeetingKey { get; set; }  // this is the unique key stored in Database to identify booking in our system, you can ignore this for POC; >0 existing meeting
        [DataMember]
        public int CancelStatus { get; set; } // to tell whether booking is in waiting or bboked it can have value like 0,1, 2,3, you can ignore this for POC; ==1 => delete from O365; ==2 - ask for acceptance; ==3 - remove from O365
        [DataMember]
        public string MailBoxOwnerEmail { get; set; }      // this is smtp address of the organiser of appointment
        [DataMember]
        public string ApiURL { get; set; } // endpoint of our system.
        [DataMember]
        public string Subject { get; set; } // subject of appointment
        [DataMember]
        public string Location { get; set; } // location of appointment
        [DataMember]
        public string StartUTC { get; set; } // strat time of appointment in utc
        [DataMember]
        public string EndUTC { get; set; } // end time of appointment in utc
        [DataMember]
        public string BookingRef { get; set; } //  custom property saved in appointment, Same is  also saved in database which allows us to link appointment with our systen
        [DataMember]
        public string SiteMailBox { get; set; } // smtp address of resource mail box
    }
}
