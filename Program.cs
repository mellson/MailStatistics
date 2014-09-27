using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using MimeKit;

namespace MailStatistics
{
    internal class Program
    {
        private const string Email = ""; // replace this with email address you are interested in getting statistics for
        private const string PathToMbox = @""; // replace this with a path to a .mbox file

        #region Fields
        static readonly IEnumerable<MimeMessage> Emails = GetMessagesFromMboxFile(PathToMbox).AsParallel();
        private static int _counter = 1;
        private static int _countTo;
        private static string _replyString;
        private static string _statusString;
        static readonly Stopwatch Stopwatch = new Stopwatch();
        #endregion Fields

        /// <summary>
        /// Main application you can run.
        /// </summary>
        private static void Main()
        {
            // Check that the user has set the right settings
            if (Email.Length == 0 || PathToMbox.Length == 0)
            {
                Logger.Out("You need to fill in the Email and the Path to your mbox file.", false);
                Logger.Out("Press any key to close the app.", false);
                Console.ReadKey();
                return;
            }

            // Write to console which file is being processed
            Logger.Out("Processing " + PathToMbox, writeToFile: false);

            // Start execution timer
            Stopwatch.Start();

            // Set the current status
            _statusString = String.Format("There are {0} emails in the mbox file", Emails.Count());
            Logger.Out(_statusString);

            // Count the number of replies
            var replies = Emails.Where(MailReplyFromMe).ToArray();
            _replyString = String.Format("{0} of these emails are replies sent by {1}", replies.Length, Email);
            Logger.Out(_replyString);
            _countTo = replies.Length;

            // Calculate statistics on reply times
            var replyTimes = replies.AsParallel().Select(TimeForReply).ToArray();
            if (replyTimes.Any())
            {
                Logger.Out("Timeformat is 'days - hours - minutes - seconds'");
                Logger.Out(String.Format("Fastest time to reply {0}", TimePrinter(replyTimes.Min())));
                Logger.Out(String.Format("Average time to reply {0}", TimePrinter(TimeSpan.FromMilliseconds(replyTimes.Select(ts => ts.TotalMilliseconds).Average()))));
                Logger.Out(String.Format("Longest time to reply {0}", TimePrinter(replyTimes.Max())));
            }
            Logger.Out("");

            // Print most emails to and from statistics
            var mostEmailsTo = Emails.ToLookup(message => message.To.ToString()).OrderByDescending(m => m.Count());
            var mostEmailsFrom = Emails.ToLookup(message => message.From.ToString()).OrderByDescending(m => m.Count());            
            PrintTopEmailers(mostEmailsTo, "sent to");
            PrintTopEmailers(mostEmailsFrom, "received from");

            // Stop the execution timer and print final info
            Stopwatch.Stop();
            Logger.Out("");
            Logger.Out(String.Format("Execution time {0}", TimePrinter(TimeSpan.FromMilliseconds(Stopwatch.ElapsedMilliseconds))));
            Console.Out.WriteLine("Press any key to close this application - results are in MailStatistics.txt");
            Console.ReadKey();
            Logger.WriteFile();
        }

        /// <summary>
        /// Prints the 5 most emailed addresses.
        /// </summary>
        /// <param name="topMailers">List of most emailed addresses.</param>
        /// <param name="text">text which indicates if the email was sent or received.</param>
        private static void PrintTopEmailers(IEnumerable<IGrouping<string, MimeMessage>> topMailers, string text)
        {
            foreach (var mail in topMailers.Take(5))
                Logger.Out(String.Format("{0,8} emails {2} {1}", mail.Count(), mail.Key, text));
        }

        /// <summary>
        /// String formatter for a TimeSpan.
        /// </summary>
        /// <param name="ts">The timespan you wish to turn into a string.</param>
        /// <returns>a string with the time.</returns>
        private static string TimePrinter(TimeSpan ts)
        {
            return ts.ToString(@"dd\.hh\:mm\:ss");
        }

        /// <summary>
        /// Calculates the time it took for this mail to be sent.
        /// </summary>
        /// <param name="replyMail">A mail identied as being a reply.</param>
        /// <returns>The time from the original mail arrived until this reply was sent.</returns>
        private static TimeSpan TimeForReply(MimeMessage replyMail)
        {
            var originalMails = Emails.Where(mail => !string.IsNullOrEmpty(mail.MessageId) && mail.MessageId.Equals(replyMail.InReplyTo));
            if (!originalMails.Any()) return new TimeSpan();
            UpdateProgress();
            return replyMail.Date - originalMails.First().Date;
        }

        /// <summary>
        /// Updates the console while running.
        /// The application is running multithreaded so the output can look weird at times.
        /// </summary>
        private static void UpdateProgress()
        {
            Console.Clear();
            Console.Out.WriteLine(_statusString);
            Console.Out.WriteLine(_replyString);
            Console.Out.WriteLine("Processing reply {0} out of {1} replies", _counter++, _countTo);
        }

        /// <summary>
        /// Checks if the mail is a reply from the Email set in the application.
        /// </summary>
        /// <param name="msg">mail you wish to check.</param>
        /// <returns>true if the mail is from the Email set in the application.</returns>
        private static bool MailReplyFromMe(MimeMessage msg)
        {
            return !string.IsNullOrEmpty(msg.InReplyTo) && msg.From.ToString().Contains(Email) && !SubjectIndicatesForward(msg);
        }

        /// <summary>
        /// Checks if the mail is a forwarded email.
        /// </summary>
        /// <param name="msg">mail you wish to check.</param>
        /// <returns>true if the mail subject indicates a forwarded message.</returns>
        private static bool SubjectIndicatesForward(MimeMessage msg)
        {
            return !string.IsNullOrEmpty(msg.Subject) && msg.Subject.ToLower().Contains("fwd:");
        }

        /// <summary>
        /// Extract messages from a mbox file.
        /// </summary>
        /// <param name="mboxPath">path to the mbox file you want to extract messages from.</param>
        /// <returns>mail messages from the mbox as MimeMessages.</returns>
        private static IEnumerable<MimeMessage> GetMessagesFromMboxFile(string mboxPath)
        {
            var parser = new MimeParser(File.OpenRead(mboxPath), MimeFormat.Mbox);
            while (!parser.IsEndOfStream)
                yield return parser.ParseMessage();
        }

        /// <summary>
        /// Simple logger which can print to console and write to a file.
        /// </summary>
        public static class Logger
        {
            public static StringBuilder LogString = new StringBuilder();
            public static void Out(string str, bool writeToFile = true)
            {
                Console.WriteLine(str);
                LogString.Append(str).Append(Environment.NewLine);
            }

            public static void WriteFile()
            {
                File.WriteAllText(@"MailsStatistics.txt", LogString.ToString());
            } 
        }
    }
}
