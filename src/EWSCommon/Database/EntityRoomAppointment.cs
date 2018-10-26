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
        public EntityRoomAppointment()
        {
            IsDeleted = false;
            IsRecurringMeeting = false;
            ExistsInExchange = false;
            DeletedLocally = false;
        }

        [Key()]
        public int Id { get; set; }

        [MaxLength(255)]
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
        /// EXCHANGE appointment unique Id
        /// </summary>
        [MaxLength(2048)]
        public string BookingId { get; set; }

        /// <summary>
        /// EXCHANGE appointment change key
        /// </summary>
        [MaxLength(2048)]
        public string BookingChangeKey { get; set; }

        /// <summary>
        /// custom property saved in appointment, Same is  also saved in database which allows us to link appointment with our systen
        /// </summary>
        [MaxLength(50)]
        public string BookingReference { get; set; }


        [MaxLength(4096)]
        public string RecurrencePattern { get; set; }


        public bool IsRecurringMeeting { get; set; }


        public bool ExistsInExchange { get; set; }


        public bool SyncedWithExchange { get; set; }


        public bool IsDeleted { get; set; }


        public bool DeletedLocally { get; set; }


        public DateTime? ModifiedDate { get; set; }
    }
}
