using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    /// <summary>
    /// Enables deep sort and distinct operations for event handling
    /// </summary>
    public class ItemEventComparer : IEqualityComparer<ItemEvent>
    {
        public bool Equals(ItemEvent x, ItemEvent y)
        {
            // If reference same object including null then return true
            if (object.ReferenceEquals(x, y))
            {
                return true;
            }

            // If one object null the return false
            if (object.ReferenceEquals(x, null) || object.ReferenceEquals(y, null))
            {
                return false;
            }

            // Compare properties for equality
            return (x.ItemId == y.ItemId && x.EventType == y.EventType
);
        }

        public int GetHashCode(ItemEvent obj)
        {
            if (object.ReferenceEquals(obj, null))
            {
                return 0;
            }

            int EventTypeHash = obj.EventType.GetHashCode();
            int ItemIdHash = obj.ItemId.GetHashCode();

            return EventTypeHash ^ ItemIdHash;
        }
    }
}
