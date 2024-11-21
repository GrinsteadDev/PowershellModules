using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
    public class AdvancedSearchEventArgs : EventArgs
    {
        public Outlook.Search SearchObject { get; private set; }
        public AdvancedSearchEventArgs(Outlook.Search search)
        {
            this.SearchObject = search;
        }
    }
    public class AttachmentContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.AttachmentSelection Attachments { get; private set; }
        public AttachmentContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.AttachmentSelection attachments)
        {
            this.CommandBar = commandBar;
            this.Attachments = attachments;
        }
    }
    public class BeforeFolderSharingDialogEventArg : EventArgs
    {
        public Outlook.MAPIFolder FolderToShare { get; private set; }
        public bool Cancel { get; set; }
        public BeforeFolderSharingDialogEventArg(Outlook.MAPIFolder folderToShare, bool cancel)
        {
            this.FolderToShare = folderToShare;
            this.Cancel = cancel;
        }
    }
    public class ContextMenuEventArgs : EventArgs
    {
        public Outlook.OlContextMenu ContextMenu { get; private set; }
        public ContextMenuEventArgs(Outlook.OlContextMenu contextMenu)
        {
            this.ContextMenu = contextMenu;
        }
    }
    public class FolderContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.MAPIFolder Folder { get; private set; }
        public FolderContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.MAPIFolder folder)
        {
            this.CommandBar = commandBar;
            this.Folder = folder;
        }
    }
    public class ItemContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.Selection Selection { get; private set; }
        public ItemContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.Selection selection)
        {
            this.CommandBar = commandBar;
            this.Selection = selection;
        }
    }
    public class ItemLoadEventArgs : EventArgs
    {
        public object Item { get; private set; }
        public ItemLoadEventArgs(object item)
        {
            this.Item = item;
        }
    }
    public class ItemSendEventArgs : EventArgs
    {
        public object Item { get; private set; }
        public bool Cancel { get; set; }
        public ItemSendEventArgs(object item, bool cancel)
        {
            this.Item = item;
            this.Cancel = cancel;
        }
    }
    public class NewMailExEventArgs : EventArgs
    {
        public string EntryID { get; private set; }
        public NewMailExEventArgs(string entryId)
        {
            this.EntryID = entryId;
        }
    }
    public class OptionsPagesAddEventArgs : EventArgs
    {
        public Outlook.PropertyPages Pages { get; private set; }
        public OptionsPagesAddEventArgs(Outlook.PropertyPages pages)
        {
            this.Pages = pages;
        }
    }
    public class ReminderEventArgs : EventArgs
    {
        public object Item { get; private set; }
        public ReminderEventArgs(object item)
        {
            this.Item = item;
        }
    }
    public class ShortcutContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.OutlookBarShortcut Shortcut { get; private set; }
        public ShortcutContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.OutlookBarShortcut shortcut)
        {
            this.CommandBar = commandBar;
            this.Shortcut = shortcut;
        }
    }
    public class StoreContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.Store Store { get; private set; }
        public StoreContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.Store store)
        {
            this.CommandBar = commandBar;
            this.Store = store;
        }
    }
    public class ViewContextMenuDisplayEventArgs : EventArgs
    {
        public Microsoft.Office.Core.CommandBar CommandBar { get; private set; }
        public Outlook.View View { get; private set; }
        public ViewContextMenuDisplayEventArgs(Microsoft.Office.Core.CommandBar commandBar, Outlook.View view)
        {
            this.CommandBar = commandBar;
            this.View = view;
        }
    }
    public class ApplicationEvent
    {
        public Outlook.Application App { get; private set; }
        public ApplicationEvent(PSObject application)
        {
            if (application.BaseObject is Outlook.Application)
            {
                this.App = (Outlook.Application)application.BaseObject;
            } else {
                throw new ArgumentException("Object must be of type " + this.App.GetType().FullName, "application");
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler> advancedSearchComplete = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler>();
        public void Add_AdvancedSearchComplete(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler del = (Outlook.Search searchObject) =>
            {
                AdvancedSearchEventArgs args = new AdvancedSearchEventArgs(searchObject);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});
            };

            this.advancedSearchComplete.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).AdvancedSearchComplete += del;
        }
        public void Remove_AdvancedSearchComplete(ScriptBlock sb)
        {
            if (this.advancedSearchComplete.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler del = this.advancedSearchComplete[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).AdvancedSearchComplete -= del;
                this.advancedSearchComplete.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AdvancedSearchStoppedEventHandler> advancedSearchStopped = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AdvancedSearchStoppedEventHandler>();
        public void Add_AdvancedSearchStopped(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_AdvancedSearchStoppedEventHandler del = (Outlook.Search searchObject) => {
                AdvancedSearchEventArgs args = new AdvancedSearchEventArgs(searchObject);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});
            };

            this.advancedSearchStopped.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).AdvancedSearchStopped += del;
        }
        public void Remove_AdvancedSearchStopped(ScriptBlock sb)
        {
            if (this.advancedSearchStopped.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_AdvancedSearchStoppedEventHandler del = this.advancedSearchStopped[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).AdvancedSearchStopped -= del;
                this.advancedSearchStopped.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler> attachmentContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler>();
        public void Add_AttachmentContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.AttachmentSelection attachments) => {
                AttachmentContextMenuDisplayEventArgs args = new AttachmentContextMenuDisplayEventArgs(commandBar, attachments);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});
            };

            this.attachmentContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).AttachmentContextMenuDisplay += del;
        }
        public void Remove_AttachmentContextMenuDisplay(ScriptBlock sb)
        {
            if (this.attachmentContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler del = this.attachmentContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).AttachmentContextMenuDisplay -= del;
                this.attachmentContextMenuDisplay.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_BeforeFolderSharingDialogEventHandler> beforeFolderSharingDialog = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_BeforeFolderSharingDialogEventHandler>();
        public void Add_BeforeFolderSharingDialog(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_BeforeFolderSharingDialogEventHandler del = (Outlook.MAPIFolder folderToShare, ref bool cancel) => {
                BeforeFolderSharingDialogEventArg args = new BeforeFolderSharingDialogEventArg(folderToShare, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});

                cancel = args.Cancel;
            };

            this.beforeFolderSharingDialog.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).BeforeFolderSharingDialog += del;
        }
        public void Remove_BeforeFolderSharingDialog(ScriptBlock sb)
        {
            if (this.beforeFolderSharingDialog.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_BeforeFolderSharingDialogEventHandler del = this.beforeFolderSharingDialog[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).BeforeFolderSharingDialog -= del;
                this.beforeFolderSharingDialog.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler> contextMenuClose = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler>();
        public void Add_ContextMenuClose(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler del = (Outlook.OlContextMenu contextMenu) => {
                ContextMenuEventArgs args = new ContextMenuEventArgs(contextMenu);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});
            };

            this.contextMenuClose.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ContextMenuClose += del;
        }
        public void Remove_ContextMenuClose(ScriptBlock sb)
        {
            if (this.contextMenuClose.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler del = this.contextMenuClose[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ContextMenuClose -= del;
                this.contextMenuClose.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler> folderContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler>();
        public void Add_FolderContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.MAPIFolder folder) => {
                FolderContextMenuDisplayEventArgs args = new FolderContextMenuDisplayEventArgs(commandBar, folder);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[] {args});
            };

            this.folderContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).FolderContextMenuDisplay += del;
        }
        public void Remove_FolderContextMenuDisplay(ScriptBlock sb)
        {
            if (this.folderContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler del = this.folderContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).FolderContextMenuDisplay -= del;
                this.folderContextMenuDisplay.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler> itemContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler>();
        public void Add_ItemContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.Selection selection) => {
                ItemContextMenuDisplayEventArgs args = new ItemContextMenuDisplayEventArgs(commandBar, selection);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.itemContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ItemContextMenuDisplay += del;
        }
        public void Remove_ItemContextMenuDisplay(ScriptBlock sb)
        {
            if (this.itemContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler del = this.itemContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ItemContextMenuDisplay -= del;
                this.itemContextMenuDisplay.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemLoadEventHandler> itemLoad = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemLoadEventHandler>();
        public void Add_ItemLoad(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ItemLoadEventHandler del = (object item) => {
                ItemLoadEventArgs args = new ItemLoadEventArgs(item);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.itemLoad.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ItemLoad += del;
        }
        public void Remove_ItemLoad(ScriptBlock sb)
        {
            if (this.itemLoad.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ItemLoadEventHandler del = this.itemLoad[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ItemLoad -= del;
                this.itemLoad.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemSendEventHandler> itemSend = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ItemSendEventHandler>();
        public void Add_ItemSend(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ItemSendEventHandler del = (object item, ref bool cancel) => {
                ItemSendEventArgs args = new ItemSendEventArgs(item, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.itemSend.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ItemSend += del;
        }
        public void Remove_ItemSend(ScriptBlock sb)
        {
            if (this.itemSend.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ItemSendEventHandler del = this.itemSend[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ItemSend -= del;
                this.itemSend.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_MAPILogonCompleteEventHandler> mapiLogonComplete = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_MAPILogonCompleteEventHandler>();
        public void Add_MAPILogonComplete(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_MAPILogonCompleteEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, null);
            };

            this.mapiLogonComplete.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).MAPILogonComplete += del;
        }
        public void Remove_MAPILogonComplete(ScriptBlock sb)
        {
            if (this.mapiLogonComplete.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_MAPILogonCompleteEventHandler del = this.mapiLogonComplete[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).MAPILogonComplete -= del;
                this.mapiLogonComplete.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_NewMailEventHandler> newMail = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_NewMailEventHandler>();
        public void Add_NewMail(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_NewMailEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, null);
            };

            this.newMail.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).NewMail += del;
        }
        public void Remove_NewMail(ScriptBlock sb)
        {
            if (this.newMail.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_NewMailEventHandler del = this.newMail[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).NewMail -= del;
                this.newMail.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_NewMailExEventHandler> newMailEx = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_NewMailExEventHandler>();
        public void Add_NewMailEx(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_NewMailExEventHandler del = (string entryId) => {
                NewMailExEventArgs args = new NewMailExEventArgs(entryId);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.newMailEx.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).NewMailEx += del;
        }
        public void Remove_NewMailEx(ScriptBlock sb)
        {
            if (this.newMailEx.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_NewMailExEventHandler del = this.newMailEx[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).NewMailEx -= del;
                this.newMailEx.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler> optionsPagesAdd = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler>();
        public void Add_OptionsPagesAdd(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler del = (Outlook.PropertyPages pages) => {
                OptionsPagesAddEventArgs args = new OptionsPagesAddEventArgs(pages);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.optionsPagesAdd.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).OptionsPagesAdd += del;
        }
        public void Remove_OptionsPagesAdd(ScriptBlock sb)
        {
            if (this.optionsPagesAdd.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler del = this.optionsPagesAdd[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).OptionsPagesAdd -= del;
                this.optionsPagesAdd.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_QuitEventHandler> quit = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_QuitEventHandler>();
        public void Add_Quit(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_QuitEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, null);
            };

            this.quit.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).Quit += del;
        }
        public void Remove_Quit(ScriptBlock sb)
        {
            if (this.quit.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_QuitEventHandler del = this.quit[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).Quit -= del;
                this.quit.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ReminderEventHandler> reminder = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ReminderEventHandler>();
        public void Add_Reminder(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ReminderEventHandler del = (object item) => {
                ReminderEventArgs args = new ReminderEventArgs(item);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.reminder.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).Reminder += del;
        }
        public void Remove_Reminder(ScriptBlock sb)
        {
            if (this.reminder.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ReminderEventHandler del = this.reminder[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).Reminder -= del;
                this.reminder.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler> shortcutContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler>();
        public void Add_ShortcutContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.OutlookBarShortcut shortcut) => {
                ShortcutContextMenuDisplayEventArgs args = new ShortcutContextMenuDisplayEventArgs(commandBar, shortcut);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.shortcutContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ShortcutContextMenuDisplay += del;
        }
        public void Remove_ShortcutContextMenuDisplay(ScriptBlock sb)
        {
            if (this.shortcutContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler del = this.shortcutContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ShortcutContextMenuDisplay -= del;
                this.shortcutContextMenuDisplay.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_StartupEventHandler> startup = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_StartupEventHandler>();
        public void Add_StartUp(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_StartupEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, null);
            };

            this.startup.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).Startup += del;
        }
        public void Remove_StartUp(ScriptBlock sb)
        {
            if (this.startup.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_StartupEventHandler del = this.startup[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).Startup -= del;
                this.startup.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_StoreContextMenuDisplayEventHandler> storeContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_StoreContextMenuDisplayEventHandler>();
        public void Add_StoreContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_StoreContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.Store store) =>
            {
                StoreContextMenuDisplayEventArgs args = new StoreContextMenuDisplayEventArgs(commandBar, store);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.storeContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).StoreContextMenuDisplay += del;
        }
        public void Remove_StoreContextMenuDisplay(ScriptBlock sb)
        {
            if (this.storeContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_StoreContextMenuDisplayEventHandler del = this.storeContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).StoreContextMenuDisplay -= del;
                this.storeContextMenuDisplay.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ViewContextMenuDisplayEventHandler> viewContextMenuDisplay = new Dictionary<ScriptBlock, Outlook.ApplicationEvents_11_ViewContextMenuDisplayEventHandler>();
        public void Add_ViewContextMenuDisplay(ScriptBlock sb)
        {
            Outlook.ApplicationEvents_11_ViewContextMenuDisplayEventHandler del = (Microsoft.Office.Core.CommandBar commandBar, Outlook.View view) => {
                ViewContextMenuDisplayEventArgs args = new ViewContextMenuDisplayEventArgs(commandBar, view);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.App) }, new object[]{args});
            };

            this.viewContextMenuDisplay.Add(sb, del);
            ((Outlook.ApplicationEvents_11_Event)this.App).ViewContextMenuDisplay += del;
        }
        public void Remove_ViewContextMenuDisplay(ScriptBlock sb)
        {
            if (this.viewContextMenuDisplay.ContainsKey(sb))
            {
                Outlook.ApplicationEvents_11_ViewContextMenuDisplayEventHandler del = this.viewContextMenuDisplay[sb];

                ((Outlook.ApplicationEvents_11_Event)this.App).ViewContextMenuDisplay -= del;
                this.viewContextMenuDisplay.Remove(sb);
            }
        }
    }
}