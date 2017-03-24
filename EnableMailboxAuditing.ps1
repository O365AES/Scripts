# This script enables Mailbox Audit Logging for a list of mailboxes.
# Instructions:
# Create "C:\temp\userlist.csv" with a list of all the email addresses of the users to be modified. Add the word "email" at the top.
# Connect to Exchange Online Powershell before running -> http://aka.ms/EXOPSPreview https://technet.microsoft.com/en-us/library/mt775114%28v=exchg.160%29.aspx


$users = Import-csv "C:\temp\userlist.csv"

foreach ($user in $users)
    {
        Set-Mailbox -Identity $user.email -AuditLogAgeLimit 90 -AuditEnabled $true -AuditAdmin Update,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf,Create,Copy,MessageBind -AuditDelegate Update,Move,MoveToDeletedItems,SoftDelete,HardDelete,FolderBind,SendAs,SendOnBehalf,Create -AuditOwner Update,MoveToDeletedItems,Move,SoftDelete,HardDelete,Create,MailboxLogin

    }