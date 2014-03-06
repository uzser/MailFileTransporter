using System;
using System.Configuration;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace MailFileTransporter
{
    internal class MailProcessor
    {
        private static readonly string FolderToSave = ConfigurationManager.AppSettings["FolderToSave"];
        private static readonly string Recipient = ConfigurationManager.AppSettings["Recipient"];
        private static readonly NameSpace OutlookNameSpace = (new Application()).GetNamespace("MAPI");
        private const string DeleteTicket = "DeleteMe";


        public void SendFilesFromFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Console.WriteLine("Directory \"{0}\" does not exist.", path);
                return;
            }

            var sentFolder = OutlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            var sentItems = sentFolder.Items;
            sentItems.ItemAdd += DeleteMarkedMailPermanently;

            foreach (var file in Directory.GetFiles(path))
            {
                try
                {
                    var message = (MailItem)(new Application()).CreateItem(OlItemType.olMailItem);
                    message.Attachments.Add(file);
                    message.Subject = DeleteTicket;
                    message.Recipients.Add(Recipient);
                    message.Send();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.WriteLine("File \"{0}\" was sent.", file);
            }
        }

        public void SetupToReceive()
        {
            var inbox = OutlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders["Personal"];
            var itemsInInbox = inbox.Items;
            itemsInInbox.ItemAdd += SaveAttachmentAndDelete;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        private void SaveAttachmentAndDelete(dynamic item)
        {
            if (item == null)
                return;

            var mail = (MailItem)item;

            try
            {
                if (!Directory.Exists(FolderToSave))
                    Directory.CreateDirectory(FolderToSave);

                var attachmentName = String.Empty;
                foreach (Attachment attachment in mail.Attachments)
                {
                    attachmentName = attachment.FileName;
                    attachment.SaveAsFile(FolderToSave + "\\" + attachmentName);
                }
                mail.Subject = DeleteTicket;

                DeleteMarkedMailPermanently(mail);
                Console.WriteLine("File \"{0}\" was saved.", attachmentName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Delete e-mail with specific mark in Subject.
        /// </summary>
        /// <param name="mail">E-mail object.</param>
        private static void DeleteMarkedMailPermanently(dynamic mail)
        {
            if (!mail.Subject.Equals(DeleteTicket))
                return;

            mail.Delete();
            mail = OutlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                .Items.Find("[Subject] = " + DeleteTicket);
            mail.Delete();
            Console.WriteLine("Mail was deleted.");
        }
    }
}