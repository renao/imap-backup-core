using System;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using System.Collections.Generic;
using MailKit.Search;
using System.IO;

namespace ImapBackup
{
    class ImapDownloader : IDisposable
    {

        ImapClient connection;

        public ImapDownloader(
            string username,
            string password,
            string host)
        {
            connection = new ImapClient();
            connection.Connect(host, 993, true);
            connection.Authenticate(username, password);
        }

        public void WriteAllMessagesToDisk(string basePath)
        {
            foreach (var folder in PersonalMailFolders)
            {
                folder.Open(FolderAccess.ReadOnly);
                int count = folder.Count;
                int progress = 0;
                Console.WriteLine("Start download for folder: " + folder.FullName + " (" + count + " Messages in Total)");
                string localFolderPath = Path.Combine(basePath, LocalFolderName(folder));
                foreach (var uid in folder.Search(SearchQuery.All))
                {
                    progress += 1;
                    Console.Write("\rSaving Mail " + progress + " from " + count);
                        SaveEmail(localFolderPath, folder.GetMessage(uid));
                }
                Console.WriteLine(folder.FullName + " done.");
            }
        }

        private void SaveEmail(string localFolderPath, MimeMessage message)
        {
            string emailFolderName = message.Date.ToString("yyyymmddhhmm");
            string emailFolderPath = Path.Combine(localFolderPath, emailFolderName);

            Directory.CreateDirectory(emailFolderPath);
            message.WriteTo(Path.Combine(emailFolderPath, "content.eml"));


            foreach (var attachment in message.Attachments)
            {
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

                if (fileName == null) continue;


                var localFileName = System.IO.Path.Combine(
                    emailFolderPath,
                    fileName);

                using (var stream = System.IO.File.Create(localFileName))
                {
                    if (attachment is MessagePart)
                    {
                        var rfc822 = (MessagePart)attachment;
                        rfc822.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)attachment;
                        part.ContentObject.DecodeTo(stream);
                    }
                }
            }

        }

        private IList<IMailFolder> PersonalMailFolders {
            get {
                return connection.GetFolders(connection.PersonalNamespaces[0]);
            }
        }

        private string LocalFolderName(IMailFolder folder)
        {
            return folder.FullName.Replace(
                folder.DirectorySeparator,
                System.IO.Path.DirectorySeparatorChar);
        }

        public void Dispose()
        {
            if (connection?.IsConnected == true)
            {
                connection.Disconnect(true);
                connection = null;
            }
        }
    }
}
