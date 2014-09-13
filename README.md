MailStatistics
==============

C# application that can look through a mbox file and gather some statistics.

Before you can use the application you need to create a mbox file.
If you use Gmail you can create a mbox file [here](https://www.google.com/settings/takeout/custom/gmail)

For this application to work you need to change two things in Program.cs

1. Set `private const string Email` to the email address of the exported mbox.

2. Set `private const string PathToMbox` to point to the path where your mbox file is located.
On windows - you copy the full path for an individual file by holding down the Shift key as you right-click the file, and then choose Copy As Path.
