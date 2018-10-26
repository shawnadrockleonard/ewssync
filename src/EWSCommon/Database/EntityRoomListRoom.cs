using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Database
{
    [Table("RoomListRooms", Schema = "dbo")]
    public class EntityRoomListRoom
    {
        [Key]
        public int Id { get; set; }

        [MaxLength(512)]
        public string Identity { get; set; }

        [MaxLength(155)]
        public string RoomList { get; set; }

        [MaxLength(155)]
        public string SmtpAddress { get; set; }


        public DateTime? LastSyncDate { get; set; }


        public int? KnownEvents { get; set; }

        /// <summary>
        /// The collection of appointments associated with the room.
        /// </summary>
        public virtual ICollection<EntityRoomAppointment> Appointments { get; set; }

        /// <summary>
        /// The collection of sync activites for the Room
        /// </summary>
        public virtual ICollection<EntityRoomListRoomSyncState> SyncStates { get; set; }
    }
}
