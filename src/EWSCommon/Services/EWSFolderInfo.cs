using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class EWSFolderInfo
    {
        public EWService Service { get; set; }


        public FolderId Folder { get; set; }


        public string SmtpAddress { get; set; }
    }
}
