using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    public class AppointmentObjectId
    {
        public Appointment Item { get; set; }

        public ItemId Id { get; set; }

        public string Base64UniqueId { get; set; }


        public string ICalUid { get; set; }

        public EmailAddress Organizer { get; set; }

        /// <summary>
        /// Represents an external source unique integer id
        /// </summary>
        public string ReferenceId { get; set; }

        /// <summary>
        /// Represents a unique character string for the meeting
        /// </summary>
        public int? MeetingKey { get; set; }
    }
}
