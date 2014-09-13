MailStatistics
==============

C# application that can look through a mbox file and gather some statistics.

# Setup

Before you can use the application you need to create a mbox file.
If you use Gmail you can create a mbox file [here](https://www.google.com/settings/takeout/custom/gmail)

For this application to work you need to change two things in Program.cs

1. Set `private const string Email` to the email address of the exported mbox.
2. Set `private const string PathToMbox` to point to the path where your mbox file is located.
On windows - you copy the full path for an individual file by holding down the Shift key as you right-click the file, and then choose Copy As Path.

# Statistics
The application will by default print the following statistics.

1. Number of emails in mbox.
2. How many of these emails are replies by the `Email` address.
3. The fastest reply time.
4. The average reply time.
5. The longest (slowest) reply time.
6. The 5 email addresses you have sent most emails to.
7. The 5 email addresses you have received most emails from.
8. Execution time of the application.
 
The result will be stored in the file MailStatistics.txt
