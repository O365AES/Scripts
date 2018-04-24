#Requires -Version 4

<#
	.SYNOPSIS
		Get mailbox forwarding options for users in Office 365.

	.DESCRIPTION
        This script will look at mailbox forwarding options provisioned using either Outlook Options method (ForwardingSmtpAddress)
        OR using an Inbox Rule.

    .PARAMETER IgnoredDomains
        An array specifying domains that you don't want to report on.
        
    .PARAMETER NoRules
        By specifying the -NoRules switch, Inbox Rules are not loaded for each mailbox, dramatically speeding up the report. However,
        we will not find auto forward rules.

    .PARAMETER OutCSV
        Output to a CSV file

	.NOTES
		Cam Murray
		Field Engineer - Microsoft
		cam.murray@microsoft.com
		
		Last update: 12 July 2017

	.LINK
		about_functions_advanced

#>

Param(
    [CmdletBinding()]
    [Array]$IgnoredDomains,
    [Switch]$NoRules,
    [String]$OutCSV=$null
)

# Create the forward array to store forwards
$Forwards = @()

# Get the mailbox objects
Write-Host "$(Get-Date) Getting mailbox objects"
$Mailboxes = (Get-Mailbox -ResultSize:Unlimited -RecipientTypeDetails UserMailbox | Select-Object Identity,ForwardingSmtpAddress)

# Loop mailboxes

$CurrentMailbox = 0

ForEach($Mailbox in $Mailboxes) {

    $CurrentMailbox++;
    If($Mailboxes.Count -gt 1) {
        Write-Progress -Activity "Checking mailbox forwards" -Status "Mailbox $($Mailbox.Name) $CurrentMailbox of $($Mailboxes.Count)" -PercentComplete ($CurrentMailbox/$Mailboxes.Count*100)
    }

    # ForwardingSmtpAddress - which can be set by the user in their Outlook Options using the default Role Assignment Policy
    If($Mailbox.ForwardingSmtpAddress -ne $null) {
        # Determine if the domain is ignored
        $Ignored = $false
        ForEach($IgnoredDomain in $IgnoredDomains) {
            If($Mailbox.ForwardingSmtpAddress -like "*$($IgnoredDomain)") { 
                $Ignored=$true
            }
        }

        # If not ignored, add to the main array
        If($Ignored -eq $false) {
            $Forwards += New-Object -TypeName psobject -Property @{
                Owner=$($Mailbox.Identity)
                Type="ForwardingSmtpAddress"
                Email=$($Mailbox.ForwardingSmtpAddress)
            }
        }

    }

    # Inbox Rules
    If(!$NoRules) {
        $Rules = Get-InboxRule -Mailbox $Mailbox.Identity | Where-Object {$_.ForwardTo -ne $null}

        ForEach($Rule in $Rules) {

            $Rule.ForwardTo | ForEach-Object {
                # Obtain the email address from the rule
                $Email = (($_ -replace """","") -split " ")[0]

                # Determine if the domain is ignored
                $Ignored = $false
                ForEach($IgnoredDomain in $IgnoredDomains) {
                    If($Email -like "*$($IgnoredDomain)") { 
                        $Ignored=$true
                    }
                }

                # If not ignored, add to the main array
                If($Ignored -eq $false) {
                    $Forwards += New-Object -TypeName psobject -Property @{
                        Owner=$($Mailbox.Identity)
                        Type="Rule"
                        Email=$($Email)
                    }
                }
            }

        }
    }
}

Write-Host "$(Get-Date) Finished. Found $($Forwards.Count) forwards in $($Mailboxes.Count) mailboxes"

$Forwards | Format-Table -AutoSize Owner,Type,Email

If($OutCSV -ne $null) {
    Write-Host "$(Get-Date) Outputting to CSV file $OutCSV"
    $Forwards | Export-CSV -Path $OutCSV -NoTypeInformation    
}