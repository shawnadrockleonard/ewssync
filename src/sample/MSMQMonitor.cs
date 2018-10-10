using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Messaging;
using ServerAgentConsole.Models;
using System.Collections.Concurrent;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;

namespace ServerAgentConsole
{
    public class MSMQMonitor
    {
        private MessageQueue mq = null;
        public ConcurrentQueue<PendingBooking> pendingBookingsQueue = new ConcurrentQueue<PendingBooking>();
        public ConcurrentQueue<UpdatedBooking> updatedBookingsQueue = new ConcurrentQueue<UpdatedBooking>();
        public ConcurrentQueue<BlockBooking> blockBookingsQueue = new ConcurrentQueue<BlockBooking>();
        public ConcurrentQueue<ChangedAppointment> changedAppointmentQueue = new ConcurrentQueue<ChangedAppointment>();
        public ConcurrentQueue<CancelledAppointment> cancelledAppointmentQueue = new ConcurrentQueue<CancelledAppointment>();
        public List<UpdatedBooking> UpdatedBookingList = new List<UpdatedBooking>();
        public List<PendingBooking> PendingBookingList = new List<PendingBooking>();
        public List<BlockBooking> blockBookingList = new List<BlockBooking>();

        private static Logger logger = LogManager.GetCurrentClassLogger();
        public MSMQMonitor(SettingsModel settings)
        {
            string queuePath = @".\Private$\"+ settings.clientName;
            try
            {
                // Open or create message queue
                if (MessageQueue.Exists(queuePath))
                    mq = new MessageQueue(queuePath);
                else
                    mq = MessageQueue.Create(queuePath);
            }
            catch (Exception ex)
            {
                logger.Error("Create/Open message queue : Exception = " + ex.Message);             
            }
            if (mq != null)
            {
                mq.ReceiveCompleted += new ReceiveCompletedEventHandler(MyReceiveCompleted);
                mq.BeginReceive();                          
            }
        }
        private  void MyReceiveCompleted(Object source, ReceiveCompletedEventArgs asyncResult)
        {
            // Connect to the queue.
            mq = (MessageQueue)source;
            try
            {                
                Message msg = mq.EndReceive(asyncResult.AsyncResult);
                PendingBooking pb = null;
                UpdatedBooking ub = null;
                BlockBooking bb = null;
                ChangedAppointment ca = null;
                CancelledAppointment cancelledAppointment = null;
                string label = (string)msg.Label;
                logger.Info("label created fine"); //temp changes should be removed
                if (label == "pending")
                {
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(PendingBooking) });
                    logger.Info("pending formatter  created fine");//temp changes should be removed
                    pb = (PendingBooking)msg.Body;
                    logger.Info("#STPendingSAR# " + pb.MeetingKey + " " + DateTime.UtcNow.Ticks);
                    logger.Info("pendingbookin conversion  created fine");//temp changes should be removed
                                                                          // Display the message information on the screen.
                    logger.Info("pendingbooking meeting key" + pb.MeetingKey);//temp changes should be removed
                    logger.Info("PendingBookingList Status " + PendingBookingList);//temp changes should be removed            
                    logger.Info("#STPendingSAEN# " + pb.MeetingKey + " " + DateTime.UtcNow.Ticks);
                    EnqueuePendingBooking(pb);
                }
                else if (label == "update")
                {
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(UpdatedBooking) });
                    logger.Info("update formatter  created fine");//temp changes should be removed
                    ub = (UpdatedBooking)msg.Body;
                    logger.Info("update conversion  created fine");//temp changes should be removed
                                                                   // Display the message information on the screen.
                    EnqueueUpdatedMSSBooking(ub);
                }
                else if (label == "Create Appointment" || label == "Update Appointment")
                {
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(ChangedAppointment) });
                    logger.Info("Appointment formatter  created fine");//temp changes should be removed
                    ca = (ChangedAppointment)msg.Body;
                    logger.Info("Appointment conversion  created fine");//temp changes should be removed

                    EnqueueChangedAppointment(ca);
                }
                else if (label == "Delete Appointment") {
                    logger.Info("start reading deleted appointment");
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(CancelledAppointment) });
                    logger.Info("Appointment formatter  created fine");
                    cancelledAppointment = (CancelledAppointment)msg.Body;
                    logger.Info("Appointment conversion  created fine");

                    EnqueueCancelledAppointment(cancelledAppointment);
                }
                else if (label == "blockbooking")
                {
                    msg.Formatter = new XmlMessageFormatter(new Type[] { typeof(BlockBooking) });
                    logger.Info("Block booking formatter created successfully"); 
                    bb = (BlockBooking)msg.Body;
                    logger.Info("Block booking conversion was successful"); //temp log
                    EnqueueBlockBooking(bb);
                }
            }
            catch (Exception ex)
            {
                logger.Error("Error in ReceiveComplete Event : Block Booking - Exception = " + ex.StackTrace + " " + ex.Message);
            }           
            mq.BeginReceive();
        }
        public void EnqueuePendingBooking(PendingBooking pendingBooking)
        {
            if (!pendingBookingsQueue.Any(x => x != null && x.MeetingKey == pendingBooking.MeetingKey))
            {
                pendingBookingsQueue.Enqueue(pendingBooking);
                logger.Info("Pending appointment added to its concurrent queue successfully");
            }
        }
        public void EnqueueUpdatedMSSBooking(UpdatedBooking updatedBooking)
        {
            if (!updatedBookingsQueue.Any(x => x != null && x.MeetingKey == updatedBooking.MeetingKey))
            {
                updatedBookingsQueue.Enqueue(updatedBooking);
                logger.Info("Updated booking added to its concurrent queue successfully");
            }
        }
        public void EnqueueChangedAppointment(ChangedAppointment changedAppointment)
        {
            if (!changedAppointmentQueue.Any(x => x != null && x.OrganizerAddress == changedAppointment.OrganizerAddress))
            {
                changedAppointmentQueue.Enqueue(changedAppointment);
                logger.Info("Changed appointment added to its concurrent queue successfully");
            }
        }
        public void EnqueueCancelledAppointment(CancelledAppointment cancelledAppointment)
        {
            if (!cancelledAppointmentQueue.Any(x => x != null && x.AppointmentICalUid == cancelledAppointment.AppointmentICalUid && x.MeetingType == cancelledAppointment.MeetingType))
            {
                cancelledAppointmentQueue.Enqueue(cancelledAppointment);
                logger.Info("Cancelled appointment added to its concurrent queue successfully");
            }
        }
        public void EnqueueBlockBooking(BlockBooking bb)
        {
            if (!blockBookingsQueue.Any(x => x != null && x.RecurrencePattern == bb.RecurrencePattern))
            {
                blockBookingsQueue.Enqueue(bb);
                logger.Info("Block booking added to its concurrent queue successfully");
            }
        }
        

    }    
}
