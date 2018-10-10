﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace EWSResourceSync.Messages
{
    [DataContract(Name = "Booking", Namespace = "http://www.tempuri.com")]
    public class BlockBooking
    {
        [DataMember]
        public string MailBoxOwnerEmail { get; set; }
        [DataMember]
        public string StartUTC { get; set; }
        [DataMember]
        public string EndUTC { get; set; }
        [DataMember]
        public string Subject { get; set; }
        [DataMember]
        public string Location { get; set; }
        [DataMember]
        public string BookingRef { get; set; }
        [DataMember]
        public string SiteMailBox { get; set; }
        [DataMember]
        public string RecurrencePattern { get; set; }  //recurrence pattern of recuuring appointments
        //public List<BookingException> Exceptions { get; set; } // exception for recurring appointments
    }
}
