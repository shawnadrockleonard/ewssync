using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    public enum BookingChangeEnum
    {
        /// <summary>
        /// you can ignore this for POC
        /// </summary>
        Ignore = 0,

        /// <summary>
        /// delete from O365
        /// </summary>
        Delete = 1,

        /// <summary>
        /// ask for acceptance
        /// </summary>
        Acceptance = 2,

        /// <summary>
        /// remove from O365
        /// </summary>
        Remove = 3
    }
}
