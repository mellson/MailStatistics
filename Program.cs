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
        static readonly IEnumerable<MimeMessage> Emails = GetMessagesFromMboxFile(PathToMbox).AsParallel();
        private static int _counter = 1;
        private static int _countTo;
        private static string _replyString;
        private static string _statusString;

        private static void Main()
        {
            Console.Out.WriteLine("Processing " + PathToMbox);
            var sw = new Stopwatch();
            sw.Start();
            _statusString = String.Format("There are {0} emails in the mbox file", Emails.Count());
            Logger.Out(_statusString);
            var replies = Emails.Where(MailReplyFromMe).ToArray();
            _replyString = String.Format("{0} of these emails are replies sent by {1}", replies.Length, Email);
            Logger.Out(_replyString);
            _countTo = replies.Length;

            var replyTimes = replies.AsParallel().Select(TimeForReply).ToArray();
            if (replyTimes.Any())
            {
                Logger.Out("Timeformat days - hours - minutes - seconds.");
                Logger.Out(String.Format("Fastest time to reply {0}", TimePrinter(replyTimes.Min())));
                Logger.Out(String.Format("Average time to reply {0}", TimePrinter(TimeSpan.FromMilliseconds(replyTimes.Select(ts => ts.TotalMilliseconds).Average()))));
                Logger.Out(String.Format("Longest time to reply {0}", TimePrinter(replyTimes.Max())));
            }

            Logger.Out("");
            var mostEmailsTo = Emails.ToLookup(message => message.To.ToString()).OrderBy(m => m.Count());
            var mostEmailsFrom = Emails.ToLookup(message => message.From.ToString()).OrderBy(m => m.Count());
            
            PrintTopEmailers(mostEmailsTo, "sent to");
            PrintTopEmailers(mostEmailsFrom, "received from");

            sw.Stop();
            Logger.Out("");
            Logger.Out(String.Format("Execution time {0}", TimePrinter(TimeSpan.FromMilliseconds(sw.ElapsedMilliseconds))));
            Console.Out.WriteLine("Press any key to close this application - results are in MailStatistics.txt");
            Console.ReadKey();
            Logger.WriteFile();
        }

        private static void PrintTopEmailers(IEnumerable<IGrouping<string, MimeMessage>> mostEmailsTo, string text)
        {
            foreach (var mail in mostEmailsTo.Reverse().Take(5))
                Logger.Out(String.Format("{0,8} emails {2} {1}", mail.Count(), mail.Key, text));
        }

        private static string TimePrinter(TimeSpan ts)
        {
            return ts.ToString(@"dd\.hh\:mm\:ss");
        }

        private static TimeSpan TimeForReply(MimeMessage replyMail)
        {
            var originalMails = Emails.Where(mail => !string.IsNullOrEmpty(mail.MessageId) && mail.MessageId.Equals(replyMail.InReplyTo));
            if (!originalMails.Any()) return new TimeSpan();
            UpdateProgress();
            return replyMail.Date - originalMails.First().Date;
        }

        private static void UpdateProgress()
        {
            Console.Clear();
            Console.Out.WriteLine(_statusString);
            Console.Out.WriteLine(_replyString);
            Console.Out.WriteLine("Processing reply {0} out of {1} replies", _counter++, _countTo);
        }

        private static bool MailReplyFromMe(MimeMessage msg)
        {
            return !string.IsNullOrEmpty(msg.InReplyTo) && msg.From.ToString().Contains(Email) && !SubjectIndicatesForward(msg);
        }

        private static bool SubjectIndicatesForward(MimeMessage msg)
        {
            return !string.IsNullOrEmpty(msg.Subject) && msg.Subject.ToLower().Contains("fwd:");
        }

        private static IEnumerable<MimeMessage> GetMessagesFromMboxFile(string mboxPath)
        {
            var parser = new MimeParser(File.OpenRead(mboxPath), MimeFormat.Mbox);
            while (!parser.IsEndOfStream)
                yield return parser.ParseMessage();
        }

        public static class Logger
        {
            public static StringBuilder LogString = new StringBuilder();
            public static void Out(string str)
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
