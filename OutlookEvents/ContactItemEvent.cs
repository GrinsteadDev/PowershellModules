using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
   public class ContactItemEvent : ItemEvent<Outlook.ContactItem>
   {
        public ContactItemEvent(PSObject item) : base(item)
        {
            
        }
   }
}