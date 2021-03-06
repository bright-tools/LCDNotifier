﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Runtime.InteropServices;

namespace ReminderNotifier
{
    public partial class ThisAddIn
    {
        private Outlook.Reminders m_reminders;

        private Dictionary<string, uint> reminderLookup = new Dictionary<string, uint>();
        uint latestReminderNo = 0;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.Reminder += new Outlook.ApplicationEvents_11_ReminderEventHandler(ThisApplication_Reminder);

            m_reminders = this.Application.Explorers.Application.Reminders;

            m_reminders.ReminderRemove += new Outlook.ReminderCollectionEvents_ReminderRemoveEventHandler( ThisApplication_ReminderRemoved );
            this.Application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);
        }

        private List<string> GetAllAppointmentIDs()
        {
            List<string> idList = new List<string>();

            // This part was very painful to get right because of some of the quirks of interop
            //  For example, using a foreach() was yielding access violations when attempting to
            //  access the contents of Reminder.Item.xxxx

            // The following was very helpful: https://blogs.msdn.microsoft.com/mstehle/2007/12/07/oom-net-part-2-outlook-item-leaks/

            for (uint i = 1; i <= m_reminders.Count; i++ )
            {
                Outlook.Reminder reminderObj = m_reminders[i];
                object item = reminderObj.Item;
                Outlook.AppointmentItem parent = item as Outlook.AppointmentItem;

                if (parent != null)
                {
                    Outlook.AppointmentItem reminderAppt = (Outlook.AppointmentItem)parent;

                    idList.Add(reminderAppt.GlobalAppointmentID);

                    Marshal.ReleaseComObject(parent);
                    parent = null;
                }

                Marshal.ReleaseComObject(reminderObj);
                reminderObj = null;
            }

            return idList;
        }

        void Application_OptionsPagesAdd(Outlook.PropertyPages Pages)
        {
            Pages.Add(new OptionsControl(), "");
        }

        private void ThisApplication_ReminderRemoved()
        {
            // List of all of the appointment IDs in Outlook
            List<string> idList = GetAllAppointmentIDs();

            bool cont = true;

            while (cont)
            {
                cont = false;
                // Loop all the appointment IDs currently sent to the display
                foreach (string id in reminderLookup.Keys)
                {
                    if (!idList.Contains(id))
                    {
                        removeReminder(id);
                        cont = true;
                        break;
                    }
                }
            }
        }

        private void removeReminder(string id )
        {
            ReminderInterface.RemoveMessage(reminderLookup[id]);

            reminderLookup.Remove(id);
        }

        void ThisApplication_Reminder(object item)
        {
            if (item is Outlook.AppointmentItem)
            {
                Outlook.AppointmentItem reminderAppt = (Outlook.AppointmentItem)item;
                uint reminderNo;

                /* Already in our list? */
                if (reminderLookup.Keys.Contains(reminderAppt.GlobalAppointmentID))
                {
                    reminderNo = reminderLookup[reminderAppt.GlobalAppointmentID];
                }
                else
                {
                    reminderLookup[reminderAppt.GlobalAppointmentID] = latestReminderNo;
                    reminderNo = latestReminderNo;

                    latestReminderNo++;
                }

                ReminderInterface.SendMessage(reminderNo, reminderAppt.Subject, reminderAppt.Start.ToString("HH:mm") + " " + reminderAppt.Location);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
