using EWS.Common.Database;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS.Common.Services
{
    public class GroupInfoSubscriptionConnection
    {
        public Dictionary<string, GroupInfo> _groups { get; set; }

        public Mailboxes _mailboxes { get; set; }

        private IList<StreamingSubscriptionConnection> connection { get; set; }

        private ITraceListener _traceListener { get; set; }

        public GroupInfoSubscriptionConnection(ITraceListener listener)
        {
            _traceListener = listener;

            var EwsToken = System.Threading.Tasks.Task.Run(async () =>
            {
                return await EWSConstants.AcquireTokenAsync();
            });

            var oAuthCreds = new OAuthCredentials(EwsToken.Result.AccessToken);

            _groups = new Dictionary<string, GroupInfo>();
            _mailboxes = new Mailboxes(oAuthCreds, _traceListener);
        }


        public void LoadMailboxes()
        {
            var database = EWSConstants.Config.Database;

            using (EWSDbContext context = new EWSDbContext(database))
            {
                foreach (var sMailbox in context.RoomListRoomEntities.ToList())
                {
                    var addedBox = _mailboxes.AddMailbox(sMailbox.SmtpAddress);
                    if (!addedBox)
                    {
                        _traceListener.Trace("Mailbox Add", $"Failed to add SMTP {sMailbox.SmtpAddress}");
                    }

                    MailboxInfo mailboxInfo = _mailboxes.Mailbox(sMailbox.SmtpAddress);
                    if (mailboxInfo != null)
                    {
                        GroupInfo groupInfo = null;
                        if (_groups.ContainsKey(mailboxInfo.GroupName))
                        {
                            groupInfo = _groups[mailboxInfo.GroupName];
                        }
                        else
                        {
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, _traceListener);
                            _groups.Add(mailboxInfo.GroupName, groupInfo);
                        }

                        if (groupInfo.Mailboxes.Count > 199)
                        {
                            // We already have enough mailboxes in this group, so we rename it and create a new one
                            // Renaming it means that we can still put new mailboxes into the correct group based on GroupingInformation
                            int i = 1;
                            while (_groups.ContainsKey($"{groupInfo.Name}-{i}"))
                            {
                                i++;
                            }
                            _groups.Remove(groupInfo.Name);
                            _groups.Add($"{groupInfo.Name}-{i}", groupInfo);
                            groupInfo = new GroupInfo(mailboxInfo.GroupName, mailboxInfo.SMTPAddress, mailboxInfo.EwsUrl, _traceListener);
                            _groups.Add(mailboxInfo.GroupName, groupInfo);
                        }

                        groupInfo.Mailboxes.Add(sMailbox.SmtpAddress);
                    }
                }
            }
        }

        public void OpenConnections()
        {
            foreach (var _group in _groups)
            {
                _traceListener.Trace("Open Connections", $"Opening connections for {_group.Key}");

                var groupInfo = _group.Value;
                if (groupInfo.AddGroupSubscriptions())
                {
                    groupInfo.OpenSubscription();
                    _traceListener.Trace("Opened Connections", $"{groupInfo.Mailboxes.Count()} mailboxes primed for StreamingSubscriptions.");
                }

            }
        }
    }
}
