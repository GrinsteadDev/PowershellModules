using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
   public class MeetingItemEvent : ItemEvent<Outlook.MeetingItem>
   {
        public MeetingItemEvent(PSObject item) : base(item)
        {
            
        }
   }
}