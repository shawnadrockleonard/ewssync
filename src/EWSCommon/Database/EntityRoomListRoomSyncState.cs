using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Database
{
    [Table("RoomListRoomsSyncState", Schema = "dbo")]
    public class EntityRoomListRoomSyncState
    {
        [Key]
        public int Id { get; set; }


        public int RoomId { get; set; }


        public EntityRoomListRoom Room { get; set; }


        public string SyncState { get; set; }


        public DateTime? SyncTimestamp { get; set; }
    }
}
