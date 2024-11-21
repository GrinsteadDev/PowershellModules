using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
   public class DocumentItemEvent : ItemEvent<Outlook.DocumentItem>
   {
        public DocumentItemEvent(PSObject item) : base(item)
        {
            
        }
   }
}