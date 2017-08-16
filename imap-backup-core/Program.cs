using System;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using System.IO;

namespace ImapBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome. Please wait, starting IMAP-Backup");

            Console.WriteLine("Enter your IMAP-username:");
            string username = Console.ReadLine();

            Console.WriteLine("Enter your IMAP-Password:");
            string password = Console.ReadLine();

            Console.WriteLine("Enter your IMAP-Hostadress:");
            Console.WriteLine("(IMAP Backup will automatically start afterwards)");
            string host = Console.ReadLine();

            using (var downloader = new ImapDownloader(username, password, host))
            {
                downloader.WriteAllMessagesToDisk(Path.Combine(Environment.CurrentDirectory, "imap-backup"));
            }

            Console.WriteLine("Done. Press key to exit.");
            Console.ReadKey();
        }
    }
}
