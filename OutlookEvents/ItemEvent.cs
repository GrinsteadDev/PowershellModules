using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PowershellExtensions.OutlookEvents
{
    public class AttachmentEventArgs : EventArgs
    {
        public Outlook.Attachment Attachment { get; private set; }
        public AttachmentEventArgs(Outlook.Attachment attachment)
        {
            this.Attachment = attachment;
        }
    }
    public class BeforeAttachmentEventArgs : EventArgs
    {
        public Outlook.Attachment Attachment { get; private set; }
        public bool Cancel { get; set; }
        public BeforeAttachmentEventArgs(Outlook.Attachment attachment, bool cancel)
        {
            this.Attachment = attachment;
            this.Cancel = cancel;
        } 
    }
    public class BeforeAutoSaveEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public BeforeAutoSaveEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class BeforeCheckNamesEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public BeforeCheckNamesEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class BeforeDeleteEventArgs : EventArgs
    {
        public object Item { get; private set; }
        public bool Cancel { get; set; }
        public BeforeDeleteEventArgs(object item, bool cancel)
        {
            this.Item = item;
            this.Cancel = cancel;
        }
    }
    public class CloseEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public CloseEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class CustomActionEventArgs : EventArgs
    {
        public object Action { get; private set; }
        public object Response { get; private set; }
        public bool Cancel { get; set; }
        public CustomActionEventArgs(object action, object response, bool cancel)
        {
            this.Action = action;
            this.Response = response;
            this.Cancel = cancel;
        }
    }
    public class PropertyChangeEventArgs : EventArgs
    {
        public string Name { get; private set; }
        public PropertyChangeEventArgs(string name)
        {
            this.Name = name;
        }
    }
    public class ForwardEventArgs : EventArgs
    {
        public object Forward { get; private set; }
        public bool Cancel { get; set; }
        public ForwardEventArgs(object forward, bool cancel)
        {
            this.Forward = forward;
            this.Cancel = cancel;
        }
    }
    public class OpenEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public OpenEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class ReadCompleteEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public ReadCompleteEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class ReplyEventArgs : EventArgs
    {
        public object Response { get; private set; }
        public bool Cancel { get; set; }
        public ReplyEventArgs(object response, bool cancel)
        {
            this.Response = response;
            this.Cancel = cancel;
        }
    }
    public class SendEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public SendEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public class WriteEventArgs : EventArgs
    {
        public bool Cancel { get; set; }
        public WriteEventArgs(bool cancel)
        {
            this.Cancel = cancel;
        }
    }
    public abstract class ItemEvent<T>
    {
        private object _item;
        public virtual T Item {
            get {
                return (T)this._item;
            }
            set {
                this._item = value;
            }
        }
        public ItemEvent(PSObject item)
        {
            if (item.BaseObject is T)
            {
                this.Item = (T)item.BaseObject;
            } else {
                throw new ArgumentException("Object must be of type " + this.Item.GetType().FullName, "item");
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_AfterWriteEventHandler> afterWrite = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_AfterWriteEventHandler>();
        public void Add_AfterWrite(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_AfterWriteEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, null);
            };

            this.afterWrite.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).AfterWrite += del;
        }
        public void Remove_AfterWrite(ScriptBlock sb)
        {
            if (this.afterWrite.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_AfterWriteEventHandler del = this.afterWrite[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).AfterWrite -= del;
                this.afterWrite.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentAddEventHandler> attachmentAdd = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentAddEventHandler>();
        public void Add_AttachmentAdd(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_AttachmentAddEventHandler del = (Outlook.Attachment attachment) => {
                AttachmentEventArgs args = new AttachmentEventArgs(attachment);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
            };

            this.attachmentAdd.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).AttachmentAdd += del;
        }
        public void Remove_AttachmentAdd(ScriptBlock sb)
        {
            if (this.attachmentAdd.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_AttachmentAddEventHandler del = this.attachmentAdd[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).AttachmentAdd -= del;
                this.attachmentAdd.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentReadEventHandler> attachmentRead = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentReadEventHandler>();
        public void Add_AttachmentRead(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_AttachmentReadEventHandler del = (Outlook.Attachment attachment) => {
                AttachmentEventArgs args = new AttachmentEventArgs(attachment);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
            };

            this.attachmentRead.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).AttachmentRead += del;
        }
        public void Remove_AttachmentRead(ScriptBlock sb)
        {
            if (this.attachmentRead.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_AttachmentReadEventHandler del = this.attachmentRead[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).AttachmentRead -= del;
                this.attachmentRead.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentRemoveEventHandler> attachmentRemove = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_AttachmentRemoveEventHandler>();
        public void Add_AttachmentRemove(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_AttachmentRemoveEventHandler del = (Outlook.Attachment attachment) => {
                AttachmentEventArgs args = new AttachmentEventArgs(attachment);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
            };

            this.attachmentRemove.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).AttachmentRemove += del;
        }
        public void Remove_AttachmentRemove(ScriptBlock sb)
        {
            if (this.attachmentRemove.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_AttachmentRemoveEventHandler del = this.attachmentRemove[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).AttachmentRemove -= del;
                this.attachmentRemove.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler> beforeAttachmentAdd = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler>();
        public void Add_BeforeAttachmentAdd(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler del = (Outlook.Attachment attachment, ref bool cancel) => {
                BeforeAttachmentEventArgs args = new BeforeAttachmentEventArgs(attachment, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAttachmentAdd.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentAdd += del;
        }
        public void Remove_BeforeAttachmentAdd(ScriptBlock sb)
        {
            if (this.beforeAttachmentAdd.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler del = this.beforeAttachmentAdd[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentAdd -= del;
                this.beforeAttachmentAdd.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler> beforeAttachmentPreview = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler>();
        public void Add_BeforeAttachmentPreview(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler del = (Outlook.Attachment attachment, ref bool cancel) => {
                BeforeAttachmentEventArgs args = new BeforeAttachmentEventArgs(attachment, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAttachmentPreview.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentPreview += del;
        }
        public void Remove_BeforeAttachmentPreview(ScriptBlock sb)
        {
            if (this.beforeAttachmentPreview.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler del = this.beforeAttachmentPreview[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentPreview -= del;
                this.beforeAttachmentPreview.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler> beforeAttachmentRead = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler>();
        public void Add_BeforeAttachmentRead(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler del = (Outlook.Attachment attachment, ref bool cancel) => {
                BeforeAttachmentEventArgs args = new BeforeAttachmentEventArgs(attachment, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAttachmentRead.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentRead += del;
        }
        public void Remove_BeforeAttachmentRead(ScriptBlock sb)
        {
            if (this.beforeAttachmentRead.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler del = this.beforeAttachmentRead[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentRead -= del;
                this.beforeAttachmentRead.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler> beforeAttachmentSave = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler>();
        public void Add_BeforeAttachmentSave(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler del = (Outlook.Attachment attachment, ref bool cancel) => {
                BeforeAttachmentEventArgs args = new BeforeAttachmentEventArgs(attachment, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAttachmentSave.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentSave += del;
        }
        public void Remove_BeforeAttachmentSave(ScriptBlock sb)
        {
            if (this.beforeAttachmentSave.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler del = this.beforeAttachmentSave[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentSave -= del;
                this.beforeAttachmentSave.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler> beforeAttachmentWriteToTempFile = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler>();
        public void Add_BeforeAttachmentWriteToTempFile(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler del = (Outlook.Attachment attachment, ref bool cancel) => {
                BeforeAttachmentEventArgs args = new BeforeAttachmentEventArgs(attachment, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAttachmentWriteToTempFile.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentWriteToTempFile += del;
        }
        public void Remove_BeforeAttachmentWriteToTempFile(ScriptBlock sb)
        {
            if (this.beforeAttachmentWriteToTempFile.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler del = this.beforeAttachmentWriteToTempFile[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAttachmentWriteToTempFile -= del;
                this.beforeAttachmentWriteToTempFile.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAutoSaveEventHandler> beforeAutoSave = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeAutoSaveEventHandler>();
        public void Add_BeforeAutoSave(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeAutoSaveEventHandler del = (ref bool cancel) => {
                BeforeAutoSaveEventArgs args = new BeforeAutoSaveEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeAutoSave.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeAutoSave += del;
        }
        public void Remove_BeforeAutoSave(ScriptBlock sb)
        {
            if (this.beforeAutoSave.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeAutoSaveEventHandler del = this.beforeAutoSave[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeAutoSave -= del;
                this.beforeAutoSave.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeCheckNamesEventHandler> beforeCheckNames = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeCheckNamesEventHandler>();
        public void Add_BeforeCheckNames(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeCheckNamesEventHandler del = (ref bool cancel) => {
                BeforeCheckNamesEventArgs args = new BeforeCheckNamesEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeCheckNames.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeCheckNames += del;
        }
        public void Remove_BeforeCheckNames(ScriptBlock sb)
        {
            if (this.beforeCheckNames.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeCheckNamesEventHandler del = this.beforeCheckNames[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeCheckNames -= del;
                this.beforeCheckNames.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeDeleteEventHandler> beforeDelete = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeDeleteEventHandler>();
        public void Add_BeforeDelete(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeDeleteEventHandler del = (object item, ref bool cancel) => {
                BeforeDeleteEventArgs args = new BeforeDeleteEventArgs(item, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.beforeDelete.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeDelete += del;
        }
        public void Remove_BeforeDelete(ScriptBlock sb)
        {
            if (this.beforeDelete.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeDeleteEventHandler del = this.beforeDelete[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeDelete -= del;
                this.beforeDelete.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeReadEventHandler> beforeRead = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_BeforeReadEventHandler>();
        public void Add_BeforeRead(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_BeforeReadEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, null);
            };

            this.beforeRead.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).BeforeRead += del;
        }
        public void Remove_BeforeRead(ScriptBlock sb)
        {
            if (this.beforeRead.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_BeforeReadEventHandler del = this.beforeRead[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).BeforeRead -= del;
                this.beforeRead.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_CloseEventHandler> close = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_CloseEventHandler>();
        public void Add_Close(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_CloseEventHandler del = (ref bool cancel) => {
                CloseEventArgs args = new CloseEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.close.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Close += del;
        }
        public void Remove_Close(ScriptBlock sb)
        {
            if (this.close.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_CloseEventHandler del = this.close[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Close -= del;
                this.close.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_CustomActionEventHandler> customAction = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_CustomActionEventHandler>();
        public void Add_CustomAction(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_CustomActionEventHandler del = (object action, object response, ref bool cancel) => {
                CustomActionEventArgs args = new CustomActionEventArgs(action, response, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.customAction.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).CustomAction += del;
        }
        public void Remove_CustomAction(ScriptBlock sb)
        {
            if (this.customAction.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_CustomActionEventHandler del = this.customAction[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).CustomAction -= del;
                this.customAction.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_CustomPropertyChangeEventHandler> customPropertyChange = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_CustomPropertyChangeEventHandler>();
        public void Add_CustomPropertyChange(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_CustomPropertyChangeEventHandler del = (string name) => {
                PropertyChangeEventArgs args = new PropertyChangeEventArgs(name);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
            };

            this.customPropertyChange.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).CustomPropertyChange += del;
        }
        public void Remove_CustomPropertyChange(ScriptBlock sb)
        {
            if (this.customPropertyChange.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_CustomPropertyChangeEventHandler del = this.customPropertyChange[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).CustomPropertyChange -= del;
                this.customPropertyChange.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_ForwardEventHandler> forward = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_ForwardEventHandler>();
        public void Add_Forward(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_ForwardEventHandler del = (object forward, ref bool cancel) => {
                ForwardEventArgs args = new ForwardEventArgs(forward, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };

            this.forward.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Forward += del;
        }
        public void Remove_Forward(ScriptBlock sb)
        {
            if (this.forward.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_ForwardEventHandler del = this.forward[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Forward -= del;
                this.forward.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_OpenEventHandler> open = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_OpenEventHandler>();
        public void Add_Open(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_OpenEventHandler del = (ref bool cancel) => {
                OpenEventArgs args = new OpenEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.open.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Open += del;
        }
        public void Remove_Open(ScriptBlock sb)
        {
            if (this.open.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_OpenEventHandler del = this.open[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Open -= del;
                this.open.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_PropertyChangeEventHandler> propertyChange = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_PropertyChangeEventHandler>();
        public void Add_PropertyChange(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_PropertyChangeEventHandler del = (string name) => {
                PropertyChangeEventArgs args = new PropertyChangeEventArgs(name);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
            };

            this.propertyChange.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).PropertyChange += del;
        }
        public void Remove_PropertyChange(ScriptBlock sb)
        {
            if (this.propertyChange.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_PropertyChangeEventHandler del = this.propertyChange[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).PropertyChange -= del;
                this.propertyChange.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReadEventHandler> read = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReadEventHandler>();
        public void Add_Read(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_ReadEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, null);
            };

            this.read.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Read += del;
        }
        public void Remove_Read(ScriptBlock sb)
        {
            if (this.read.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_ReadEventHandler del = this.read[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Read -= del;
                this.read.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReadCompleteEventHandler> readComplete = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReadCompleteEventHandler>();
        public void Add_ReadComplete(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_ReadCompleteEventHandler del = (ref bool cancel) => {
                ReadCompleteEventArgs args = new ReadCompleteEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.readComplete.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).ReadComplete += del;
        }
        public void Remove_ReadComplete(ScriptBlock sb)
        {
            if (this.readComplete.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_ReadCompleteEventHandler del = this.readComplete[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).ReadComplete -= del;
                this.readComplete.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReplyEventHandler> reply = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReplyEventHandler>();
        public void Add_Reply(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_ReplyEventHandler del = (object response, ref bool cancel) => {
                ReplyEventArgs args = new ReplyEventArgs(response, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.reply.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Reply += del;
        }
        public void Remove_Reply(ScriptBlock sb)
        {
            if (this.reply.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_ReplyEventHandler del = this.reply[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Reply -= del;
                this.reply.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReplyAllEventHandler> replyAll = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_ReplyAllEventHandler>();
        public void Add_ReplyAll(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_ReplyAllEventHandler del = (object response, ref bool cancel) => {
                ReplyEventArgs args = new ReplyEventArgs(response, cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.replyAll.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).ReplyAll += del;
        }
        public void Remove_ReplyAll(ScriptBlock sb)
        {
            if (this.replyAll.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_ReplyAllEventHandler del = this.replyAll[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).ReplyAll -= del;
                this.replyAll.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_SendEventHandler> send = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_SendEventHandler>();
        public void Add_Send(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_SendEventHandler del = (ref bool cancel) => {
                SendEventArgs args = new SendEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.send.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Send += del;
        }
        public void Remove_Send(ScriptBlock sb)
        {
            if (this.send.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_SendEventHandler del = this.send[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Send -= del;
                this.send.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_UnloadEventHandler> unload = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_UnloadEventHandler>();
        public void Add_Unload(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_UnloadEventHandler del = () => {
                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, null);
            };

            this.unload.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Unload += del;
        }
        public void Remove_Unload(ScriptBlock sb)
        {
            if (this.unload.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_UnloadEventHandler del = this.unload[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Unload -= del;
                this.unload.Remove(sb);
            }
        }
        private Dictionary<ScriptBlock, Outlook.ItemEvents_10_WriteEventHandler> write = new Dictionary<ScriptBlock, Outlook.ItemEvents_10_WriteEventHandler>();
        public void Add_Write(ScriptBlock sb)
        {
            Outlook.ItemEvents_10_WriteEventHandler del = (ref bool cancel) => {
                WriteEventArgs args = new WriteEventArgs(cancel);

                sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});

                cancel = args.Cancel;
            };
            
            this.write.Add(sb, del);
            ((Outlook.ItemEvents_10_Event)this.Item).Write += del;
        }
        public void Remove_Write(ScriptBlock sb)
        {
            if (this.write.ContainsKey(sb))
            {
                Outlook.ItemEvents_10_WriteEventHandler del = this.write[sb];

                ((Outlook.ItemEvents_10_Event)this.Item).Write -= del;
                this.write.Remove(sb);
            }
        }
    }
    // ((Outlook.ItemEvents_10_Event)this.Item)
    // sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, new object[]{args});
    // sb.InvokeWithContext(null, new List<PSVariable>() { new PSVariable("this", this.Item) }, null);
}