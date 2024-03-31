# OutlookSendTrayNotificationScript
This PowerShell script checks if there are any items left in the default Outlook account's Sent Items folder and sends a desktop notification if items are found.

The script performs the following steps:
1. Defines a function `Show-DesktopNotification` to display desktop notifications with a specified title and message.
2. Creates an Outlook application object to interact with Outlook.
3. Retrieves the number of email items in the Sent Items folder of the default Outlook account.
4. If there are items present, it triggers a desktop notification indicating the number of unsent emails.
5. If the Sent Items folder is empty, it outputs a message indicating that the folder is empty.
6. Releases the Outlook object to free up system resources.
