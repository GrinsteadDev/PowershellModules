using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
   public class AppointmentItemEvent : ItemEvent<Outlook.AppointmentItem>
   {
        public AppointmentItemEvent(PSObject item) : base(item)
        {
            
        }
   }
}