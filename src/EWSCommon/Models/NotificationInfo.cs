using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Models
{
    public struct NotificationInfo
    {
        public string Mailbox { get; set; }

        public object Event { get; set; }

        public ExchangeService Service { get; set; }
    }
}
