using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
   public class DistListItemEvent : ItemEvent<Outlook.DistListItem>
   {
        public DistListItemEvent(PSObject item) : base(item)
        {
            
        }
   }
}