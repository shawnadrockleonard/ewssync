using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Database
{
    [Table("RoomAppointment", Schema = "dbo")]
    public class EntityRoomAppointment
    {
        [Key()]
        public int Id { get; set; }


        public string OrganizerSmtpAddress { get; set; }

        [Required]
        public int RoomId { get; set; }

        public EntityRoomListRoom Room { get; set; }


        [Required]
        public DateTime StartUTC { get; set; }

        [Required]
        public DateTime EndUTC { get; set; }

        [MaxLength(256)]
        public string Subject { get; set; }

        [MaxLength(1024)]
        public string Location { get; set; }

        /// <summary>
        /// Foreign key or unique constraint related to this appointment
        /// </summary>
        [MaxLength(50)]
        public string BookingReference { get; set; }


        [MaxLength(4096)]
        public string RecurrencePattern { get; set; }


        public bool IsRecurringMeeting { get; set; }


        public bool ExistsInExchange { get; set; }

        /// <summary>
        /// Gets the unique Id of the Exchange object.
        /// </summary>
        public string AppointmentUniqueId { get; set; }
    }
}
