# Outlook-Email-Downloader

Welcome to the Outlook-Email-Downloader wiki!

Script Capabilities:

This script can download mail attachments from a specified Outlook folder to a specified location.

Specific mails can be chosen by filtering the Email subjects using Regex.

The downloaded attachments can be converted to CSV format

The downloaded attachments can be Unzipped.

After download a specified MS SQL Job can be triggered.

Options that can be put in the arguments field in the emails table:

Move mail to 'Deleted Items' using the 'del' command.

Mark mails as unread for test purposes using the 'unr' command.

Unzip the downloaded attachments using the 'unzip' command.

Rename attachments using the 'rename' command

Rename attachments by also adding the email receive timestamp at the end of the name by using the command 'Drename'

Convert attachments to CSV using the 'csv' command.

Example of all arguments that can be included the the DB field 'args': del;unr;unzip;rename=NewName;csv;Drename=NewName
