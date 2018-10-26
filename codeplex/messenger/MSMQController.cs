using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Messaging;
using ServerProcessConsole.Models;
using NLog;

namespace ServerProcessConsole
{
    public class MSMQController
    {
        private MessageQueue mq;
        private List<MessageQueue> queueLists;
        private int bookingCounter = 0;
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public MSMQController(SettingsModel settings)
        {
            queueLists = new List<MessageQueue>();
            MessageQueue mq;
            string machinesString = settings.MachineName;
            string[] machines = machinesString.Split(',');

            for (int i = 0; i < machines.Length; ++i)
            {
                string clientName = settings.ClientName; // seperate MSMQ name for each client - queue name same as client name
                string queuePath = "FormatName:Direct=OS:" + machines[i] + "\\Private$\\" + clientName;
                logger.Info("MSMQ Path: " + queuePath);
                mq = new MessageQueue(queuePath);
                queueLists.Add(mq);
            }
        }

        /// <summary>
        /// Create/ Send messages to MSMQ
        /// </summary>
        public void SendPendingMessage(PendingBooking pb, UpdatedBooking ub, BlockBooking bb)
        {
            //int start = 0;
            int maxNumberofqueues = queueLists.Count - 1;
            Message message = new Message();
            message.Recoverable = true;
            if (pb != null)
            {
               
                message.Body = pb;
                message.Label = "pending";
            }
            else if (ub != null)
            {
               
                message.Body = ub;
                message.Label = "update";
            }    
            else if (bb != null)
            {
                message.Body = bb;
                message.Label = "blockbooking";
            }      
            mq = queueLists[bookingCounter];
            logger.Debug("message sending to " + mq.Path);
            logger.Debug("message: " + message.ToString());
            mq.Send(message);
            if (bookingCounter == maxNumberofqueues)
            {
                bookingCounter = 0;
            }
            else
            {
                bookingCounter++;
            }            
        }

    }
}
