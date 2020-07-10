using DefaultDelaySend.Properties;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DefaultDelaySend
{
    [ComVisible(true)]
    public class DefaultDelaySendRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI explorerRibbon;

        public DefaultDelaySendRibbon()
        {
        }

        private void Inspectors_NewInspector(Inspector Inspector)
        {
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate += DefaultDelaySendRibbon_Activate;
        }

        private void DefaultDelaySendRibbon_Activate()
        {
            this.explorerRibbon.Invalidate();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXML = String.Empty;

            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                ribbonXML = GetResourceText("DefaultDelaySend.MailTab.xml");
            }
            else if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                ribbonXML = GetResourceText("DefaultDelaySend.ExplorerTab.xml");
            }

            return ribbonXML;
        }

        #endregion

        #region Ribbon Callbacks

        public void ExplorerRibbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.explorerRibbon = ribbonUI;

            Globals.ThisAddIn.Application.Inspectors.NewInspector += Inspectors_NewInspector;
        }

        public void outOfOfficeDayCheckBox_onAction(Office.IRibbonControl control, bool isPressed)
        {
            if (control.Id.StartsWith("DelaySend_OutOfOffice"))
            {
                string settingName = control.Id.Replace("DelaySend_", "");

                Type defaultType = typeof(DefaultDelaySendSettings);
                PropertyInfo propertyInfo = defaultType.GetProperty(settingName);

                if (propertyInfo != null)
                {
                    propertyInfo.SetValue(DefaultDelaySendSettings.Default, isPressed);
                    DefaultDelaySendSettings.Default.Save();
                    ThisAddIn.RefreshSettings();
                }
            }
        }

        public bool outOfOfficeDayCheckBox_getPressed(Office.IRibbonControl control)
        {
            if (control.Id.StartsWith("DelaySend_OutOfOffice"))
            {
                string settingName = control.Id.Replace("DelaySend_", "");

                Type defaultType = typeof(DefaultDelaySendSettings);
                PropertyInfo propertyInfo = defaultType.GetProperty(settingName);

                if (propertyInfo != null)
                {
                    return (bool)propertyInfo.GetValue(DefaultDelaySendSettings.Default);
                }
            }

            return false;
        }

        public void businessHourEditBox_onChange(Office.IRibbonControl control, string text)
        {
            if (control.Id.StartsWith("DelaySend_"))
            {
                string settingName = control.Id.Replace("DelaySend_", "");
                int value;

                if (settingName == "StartBusinessHour")
                {
                    if (Int32.TryParse(text, out value))
                    {
                        DefaultDelaySendSettings.Default.StartBusinessHour = value;
                    }
                }
                else if (settingName == "EndBusinessHour")
                {
                    if (Int32.TryParse(text, out value))
                    {
                        DefaultDelaySendSettings.Default.EndBusinessHour = value;
                    }
                }

                DefaultDelaySendSettings.Default.Save();
                ThisAddIn.RefreshSettings();
            }
        }

        public string businessHourEditBox_getText(Office.IRibbonControl control)
        {
            if (control.Id.StartsWith("DelaySend_"))
            {
                string settingName = control.Id.Replace("DelaySend_", "");

                if (settingName == "StartBusinessHour")
                {
                    return DefaultDelaySendSettings.Default.StartBusinessHour.ToString();
                }
                else if (settingName == "EndBusinessHour")
                {
                    return DefaultDelaySendSettings.Default.EndBusinessHour.ToString();
                }
            }

            return "";
        }

        public void overrideDelaySendCheckBox_onAction(Office.IRibbonControl control, bool isPressed)
        {
            object context = control.Context;
            Func<Outlook.MailItem, bool> evaluateDelaySendOverrideLambda = (Outlook.MailItem mailItem) =>
            {
                UserProperty prop = mailItem.UserProperties.Find("OverrideDelaySend");

                if (prop == null)
                {
                    prop = mailItem.UserProperties.Add("OverrideDelaySend", OlUserPropertyType.olYesNo);
                }

                if (isPressed)
                {
                    prop.Value = true;
                }
                else
                {
                    prop.Value = false;
                }

                return true;
            };

            if (context is Outlook.Explorer)
            {
                var explorer = context as Outlook.Explorer;

                Outlook.Selection selection = null;
                try
                {
                    selection = explorer.Selection;
                }
                catch (System.Exception)
                {

                }

                if (selection != null)
                {
                    foreach (object item in selection)
                    {
                        if (item is Outlook.MailItem)
                        {
                            evaluateDelaySendOverrideLambda(item as Outlook.MailItem);
                        }
                    }
                }
            }
            else if (context is Outlook.Inspector)
            {
                var inspector = context as Outlook.Inspector;
                object item = inspector.CurrentItem;
               
                if (item is Outlook.MailItem)
                {
                    evaluateDelaySendOverrideLambda(item as Outlook.MailItem);
                }
            }
        }

        public bool overrideDelaySendCheckBox_getPressed(Office.IRibbonControl control)
        {
            return false;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
