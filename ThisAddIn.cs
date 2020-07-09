using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using DefaultDelaySend.Properties;

namespace DefaultDelaySend
{
    public partial class ThisAddIn
    {
        public static HashSet<DayOfWeek> OutOfOfficeDays = new HashSet<DayOfWeek>();

        public static void RefreshSettings()
        {
            var days = new (DayOfWeek, bool)[]
            {
                (DayOfWeek.Monday, DefaultDelaySendSettings.Default.OutOfOfficeMonday),
                (DayOfWeek.Tuesday, DefaultDelaySendSettings.Default.OutOfOfficeTuesday),
                (DayOfWeek.Wednesday, DefaultDelaySendSettings.Default.OutOfOfficeWednesday),
                (DayOfWeek.Thursday, DefaultDelaySendSettings.Default.OutOfOfficeThursday),
                (DayOfWeek.Friday, DefaultDelaySendSettings.Default.OutOfOfficeFriday),
                (DayOfWeek.Saturday, DefaultDelaySendSettings.Default.OutOfOfficeSaturday),
                (DayOfWeek.Sunday, DefaultDelaySendSettings.Default.OutOfOfficeSunday)
            };

            OutOfOfficeDays.Clear();

            foreach (var day in days)
            {
                if (day.Item2)
                {
                    OutOfOfficeDays.Add(day.Item1);
                }
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RefreshSettings();

            Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool Cancel)
        {
            // Only apply delay delivery to mail items.
            if (item is Outlook.MailItem == false)
            {
                return;
            }

            var mailItem = item as Outlook.MailItem;

            var sendDate = DateTime.Now;
            var delayDate = DateTime.Now;
            bool outsideBusinessHours = false;

            // Check to see if this mail is being sent outside of business hours.
            if (OutOfOfficeDays.Contains(sendDate.DayOfWeek) ||
                (sendDate.Hour > DefaultDelaySendSettings.Default.EndBusinessHour))
            {
                // This mail was sent on a non-business day or after business hours on a business day, so
                // we need to delay to at least the next day (and possibly longer).
                outsideBusinessHours = true;

                do
                {
                    delayDate = delayDate.AddDays(1);

                    // If all of the business days are out of office, then return so that we don't
                    // loop forever.
                    if (delayDate.DayOfWeek == sendDate.DayOfWeek)
                    {
                        return;
                    }

                } while (OutOfOfficeDays.Contains(delayDate.DayOfWeek));

                delayDate = new DateTime(delayDate.Year, delayDate.Month, delayDate.Day, DefaultDelaySendSettings.Default.StartBusinessHour, 0, 0);
            }
            else if (sendDate.Hour <= DefaultDelaySendSettings.Default.StartBusinessHour)
            {
                // This mail was sent before business hours on the current day, so delay send to the start hour.
                outsideBusinessHours = true;

                delayDate = new DateTime(delayDate.Year, delayDate.Month, delayDate.Day, DefaultDelaySendSettings.Default.StartBusinessHour, 0, 0);
            }

            if (outsideBusinessHours == false)
            {
                return;
            }

            // Check to see if we should override delay send.
            UserProperty overrideDelaySendProp = mailItem.UserProperties.Find("OverrideDelaySend");
            bool overrideDelaySend = (overrideDelaySendProp != null && overrideDelaySendProp.Value == true);

            if (overrideDelaySend == true)
            {
                mailItem.DeferredDeliveryTime = sendDate;
                return;
            }

            // At this point, the mail has been sent outside of business hours and
            // the user did not elect to override delay send. Delay delivery to the
            // next business day.

            var prop = mailItem.UserProperties.Add("DelaySend", OlUserPropertyType.olYesNo);
            prop.Value = true;

            mailItem.DeferredDeliveryTime = delayDate;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DefaultDelaySendRibbon();
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
