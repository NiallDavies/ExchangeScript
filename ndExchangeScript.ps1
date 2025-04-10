Import-Module ExchangeOnlineManagement 

# Define Functions first
Function DisplayMenu() { # Main Menu of the script

    $menuloop = $true
    while($menuloop) {
          $choice = Read-Host -Prompt "Please select one of the following options:`n
                [0]     Export list of mailboxes
                [1]     Export folder sizes for a mailbox
                [2]     Export distribution group members
                [3]     Export list of distribution groups
                [4]     Export list of Office365 group members
                [5]     Export delegate permissions list (Send As/Full Access)
                [6]     Import contacts from CSV
                [7]     Add members to distribution group from CSV
                [8]     Add user delegate permissions to mailboxes from CSV
                [9999]  Enter custom Exchange Online commands in the current session (Advanced)
                [E]     Exit script`n`n"
    
            switch ($choice) {
                0 { ExportMailboxList }
                1 { ExportFolderSize }
                2 { ExportDistListMembers }
                3 { ExportDistLists }
                4 { ExportGroupMembers }
                5 { ExportDelegatePermissions }
                6 { ImportContacts }
                7 { ImportToDistList }
                8 { ImportDelegates }
                E { $menuloop = $false 
                        ExitScript }
                9999 { $menuloop = $false }
                Default { 
                    Write-Host "`n"
                    Write-Warning "Sorry, I didn't understand that!" 
                    Write-Host "`n"
                }
            }
    }
}

Function ExportMailboxList { # Exports Mailbox list to CSV file
    Write-Host "`n"
    
    $exportloop = $true
    
    while($exportloop) {
        $Option = Read-Host -Prompt "Please select one of the following options.
            [0] Export ALL mailboxes
            [1] Export USER mailboxes only
            [2] Export SHARED mailboxes only
            [E] Return to Main Menu
            "
        switch ($Option) {
            0 { 
                Write-Host "`n"
                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\Mailbox-list.csv"

                Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, recipienttypedetails, @{Name='Mailbox Size';Expression={Get-MailboxStatistics $_.UserPrincipalName | Select-Object TotalItemSize}} | Sort-Object PrimarySmtpAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Exported mailbox list to $FilePath" -ForegroundColor Green
                $exportloop = $false
            }
            1 { 
                Write-Host "`n"
                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\Mailbox-list.csv"

                Get-Mailbox -ResultSize Unlimited  -Filter {recipienttypedetails -eq "UserMailbox"} | Select-Object DisplayName, PrimarySmtpAddress, recipienttypedetails, @{Name='Mailbox Size';Expression={Get-MailboxStatistics $_.UserPrincipalName | Select-Object TotalItemSize}} | Sort-Object PrimarySmtpAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Exported mailbox list to $FilePath" -ForegroundColor Green 
                $exportloop = $false
            }
            2 {
                Write-Host "`n"
                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\Mailbox-list.csv"

                Get-Mailbox -ResultSize Unlimited -Filter {recipienttypedetails -eq "SharedMailbox"} | Select-Object DisplayName, PrimarySmtpAddress, recipienttypedetails, @{Name='Mailbox Size';Expression={Get-MailboxStatistics $_.UserPrincipalName | Select-Object TotalItemSize}} | Sort-Object PrimarySmtpAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Exported mailbox list to $FilePath" -ForegroundColor Green 
                $exportloop = $false                
            }
            E {
                $exportloop = $false
            }
            default{ 
                Write-Host "`n"
                Write-Warning "Sorry, I didn't understand that!"
                Write-Host "`n" 
            }
        }
    }
    Write-Host "`n"
}

Function ExportFolderSize {
    Write-Host "`n"
    $Account = Read-Host -Prompt "Please enter the email address for the account you would like to export the folder sizes of"
    $exportloop = $true
    
    while($exportloop) {
        $Archive = Read-Host -Prompt "Is there also an online archive that you would like to export the folder size for? [Y/N]"
        switch ($archive) {
            Y { $archiveexport = $true
                $exportloop = $false
            }
            N { $archiveexport = $false
                $exportloop = $false
            }
            default{ 
                Write-Host "`n"
                Write-Warning "Sorry, I didn't understand that!" 
                Write-Host "`n"
            }
        }
    }
    Write-Host "Exporting information, this could take some time..."
    $FilePath = "$PSScriptRoot\Exported\$Account-Foldersize.csv"
    $FilePathArchive = "$PSScriptRoot\Exported\$Account-ArchiveFoldersize.csv"
    Get-MailboxFolderStatistics -Identity $Account | Select-Object Name, FolderPath, ItemsInFolder, FolderSize | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported folder list for $Account to $FilePath" -ForegroundColor Green
    if ($Archiveexport) {
        Get-MailboxFolderStatistics -Identity $Account -Archive | Select-Object Name, FolderPath, ItemsInFolder, FolderSize | Export-CSV $FilePathArchive -NoTypeInformation -Encoding UTF8
        Write-Host "Exported folder list for $Account archive to $FilePathArchive" -ForegroundColor Green
    }
    Write-Host "`n"
}

Function ExportDistListMembers { # Exports Distribution group members to CSV file
    Write-Host "`n"
    $exportloop = $true
    while($exportloop) {
        $Option = Read-Host -Prompt "Please select one of the following options.
            [0] Export members for ALL distribution lists
            [1] Export delegate permissions for a particular distribution list
            [E] Return to main menu
            "
        switch ($Option) { 
            1 {
                $DGName = Read-Host -Prompt "Please enter the name of the distribution list you would like to export the members of"
                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\$DGName-Members.csv"
                Get-DistributionGroupMember -Identity $DGName | Select-Object Name, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Exported members of $DGName to $FilePath" -ForegroundColor Green
                Write-Host "`n"
                $exportloop = $false
            }
            0 {
                
                Write-Host "Exporting information, this could take some time..."
                

                $DL = Get-DistributionGroup
                ForEach ($List in $DL) {
                    $DGEmail = $List.PrimarySMTPAddress
                    $DGName = $List.DisplayName
                    $FilePath = "$PSScriptRoot\Exported\$DGName-Members.csv"
                    Get-DistributionGroupMember -Identity $DGEmail | Select-Object Name, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                    Write-Host "Exported members of $DGName to $FilePath" -ForegroundColor Green
                }

               # Get-DistributionGroupMember | Select-Object Name, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Done" -ForegroundColor Green
                Write-Host "`n"
                $exportloop = $false
            }
            E {
                $exportloop = $false
            }
            default {
                Write-Host "`n"
                Write-Warning "Sorry, I didn't understand that!" 
                Write-Host "`n"                
            }
        }
    }
    
}

Function ExportGroupMembers { # Exports Office365 group members to CSV file
    Write-Host "`n"
    $exportloop = $true
    while($exportloop) {
        $Option = Read-Host -Prompt "Please select one of the following options.
            [0] Export members for ALL office365 groups
            [1] Export delegate permissions for a particular office365 group
            [E] Return to main menu
            "
        switch ($Option) { 
            1 {
                $DGName = Read-Host -Prompt "Please enter the name of the Office365 group you would like to export the members of"
                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\$DGName-GroupMembers.csv"
                Get-UnifiedGroup -Identity $DGName | Get-UnifiedGroupLinks -LinkType Member | Select-Object DisplayName, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Exported members of $DGName to $FilePath" -ForegroundColor Green
                Write-Host "`n"
                $exportloop = $false
            }
            0 {
                
                Write-Host "Exporting information, this could take some time..."
                

                $DL = Get-UnifiedGroup
                ForEach ($List in $DL) {
                    $DGEmail = $List.PrimarySMTPAddress
                    $DGName = $List.DisplayName
                    $FilePath = "$PSScriptRoot\Exported\$DGEmail-GroupMembers.csv"
                    Get-UnifiedGroup -Identity $DGName | Get-UnifiedGroupLinks -LinkType Member | Select-Object DisplayName, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                    Write-Host "Exported members of $DGName to $FilePath" -ForegroundColor Green
                }

               # Get-DistributionGroupMember | Select-Object Name, PrimarySMTPAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
                Write-Host "Done" -ForegroundColor Green
                Write-Host "`n"
                $exportloop = $false
            }
            E {
                $exportloop = $false
            }
            default {
                Write-Host "`n"
                Write-Warning "Sorry, I didn't understand that!" 
                Write-Host "`n"                
            }
        }
    }
    
}

Function ExportDistLists {
    Write-Host "`n"
    Write-Host "Exporting information, this could take some time..."
    $FilePath = "$PSScriptRoot\Exported\DistributionLists.csv"
    Get-DistributionGroup | Select-Object DisplayName,PrimarySmtpAddress | Export-CSV $FilePath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported list of distribution groups to $FilePath" -ForegroundColor Green
    Write-Host "`n"
}

Function ExportDelegatePermissions {
    Write-Host "`n"
    $exportloop = $true
    
    while($exportloop) {
        $Option = Read-Host -Prompt "Please select one of the following options.
            [0] Export delegate permissions for ALL mailboxes (CSV per mailbox)
            [1] Export delegate permissions for SHARED mailboxes (CSV per mailbox)
            [2] Export delegate permissions for a particular mailbox (Single CSV)
            [3] Export delegate permissions for ALL mailboxes (Single CSV)
            [4] Export delegate permissions for SHARED mailboxes (Single CSV)
            [E] Return to main menu
            "
        switch ($Option) {
            0 {  
                Write-Host "`n"
                Write-Host "Exporting information, this could take some time..."
                

                $mball = Get-Mailbox -resultsize unlimited
                ForEach($mb in $mball) {
                    $account = $mb.PrimarySMTPAddress
                    $accountname = $mb.DisplayName
                
                    $FilePath = "$PSScriptRoot\Exported\$account-delegates-list.csv"
                    $FilePath2 = "$PSScriptRoot\Exported\$account-SendAs-list.csv"

                    Get-Mailbox -identity $Account | Get-MailboxPermission | Select-Object Identity,User,AccessRights  | Where-Object {($_.user -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath -NoTypeInformation
                    Write-Host "Exported delegate permissions list for $accountname to $FilePath" -ForegroundColor Green

                    Get-Mailbox -identity $Account | Get-RecipientPermission | select-object Identity,Trustee,AccessRights | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath2 -NoTypeInformation
                    Write-Host "Exported send as permissions list for $accountname to $FilePath2" -ForegroundColor Green
                }

                $exportloop = $false
            }  
            1 {
            
                Write-Host "`n"
                Write-Host "Exporting information, this could take some time..."

                $mball = Get-Mailbox -resultsize unlimited -Filter {recipienttypedetails -eq "SharedMailbox"}
                ForEach($mb in $mball) {
                    $account = $mb.PrimarySMTPAddress
                    $accountname = $mb.DisplayName
                
                    $FilePath = "$PSScriptRoot\Exported\$accountname-delegates-list.csv"
                    $FilePath2 = "$PSScriptRoot\Exported\$accountname-SendAs-list.csv"

                    Get-Mailbox -identity $Account | Get-MailboxPermission | Select-Object Identity,User,AccessRights  | Where-Object {($_.user -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath -NoTypeInformation
                    Write-Host "Exported delegate permissions list for $accountname to $FilePath" -ForegroundColor Green

                    Get-Mailbox -identity $Account | Get-RecipientPermission | select-object Identity,Trustee,AccessRights | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath2 -NoTypeInformation
                    Write-Host "Exported send as permissions list for $accountname to $FilePath2" -ForegroundColor Green
                }

                $exportloop = $false
            }     
            2 {

                Write-Host "`n"
                $Account = Read-Host -Prompt "Please enter the email address for the mailbox you would like to export the delegates of"
                Write-Host "`n"

                Write-Host "Exporting information, this could take some time..."
                $FilePath = "$PSScriptRoot\Exported\$account-delegates-list.csv"
                $FilePath2 = "$PSScriptRoot\Exported\$account-SendAs-list.csv"

                Get-Mailbox -identity $Account | Get-MailboxPermission | Select-Object Identity,User,AccessRights  | Where-Object {($_.user -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath -NoTypeInformation
                Write-Host "Exported delegate permissions list to $FilePath" -ForegroundColor Green

                Get-Mailbox -identity $Account | Get-RecipientPermission | select-object Identity,Trustee,AccessRights | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath2 -NoTypeInformation
                Write-Host "Exported send as permissions list to $FilePath2" -ForegroundColor Green

                $exportloop = $false
            }
            3 {
                $FilePath = "$PSScriptRoot\Exported\All-delegates-list.csv"
                $FilePath2 = "$PSScriptRoot\Exported\All-SendAs-list.csv"

                Get-Mailbox -resultsize unlimited | Get-MailboxPermission | Select-Object Identity,User,AccessRights  | Where-Object {($_.user -ne "NT AUTHORITY\SELF")} |  Export-Csv -Path $FilePath -NoTypeInformation
                Write-Host "Exported delegate permissions list for ALL MAILBOXES to $FilePath" -ForegroundColor Green
                
                Get-Mailbox -resultsize unlimited | Get-RecipientPermission | select-object Identity,Trustee,AccessRights | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath2 -NoTypeInformation
                Write-Host "Exported send as permissions list for ALL MAILBOXES to $FilePath2" -ForegroundColor Green
                $exportloop = $false
            }
            4 {
                $FilePath = "$PSScriptRoot\Exported\Sharedmailbox-delegates-list.csv"
                $FilePath2 = "$PSScriptRoot\Exported\Sharedmailbox-SendAs-list.csv"

                Get-Mailbox -resultsize unlimited -Filter {recipienttypedetails -eq "SharedMailbox"} | Get-MailboxPermission | Select-Object Identity,User,AccessRights  | Where-Object {($_.user -ne "NT AUTHORITY\SELF")} |  Export-Csv -Path $FilePath -NoTypeInformation
                Write-Host "Exported delegate permissions list for SHARED MAILBOXES to $FilePath" -ForegroundColor Green
                
                Get-Mailbox -resultsize unlimited -Filter {recipienttypedetails -eq "SharedMailbox"} | Get-RecipientPermission | select-object Identity,Trustee,AccessRights | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")} | Export-Csv -Path $FilePath2 -NoTypeInformation
                Write-Host "Exported send as permissions list for SHARED MAILBOXES to $FilePath2" -ForegroundColor Green
                $exportloop = $false
            }
            E {
                $exportloop = $false
            }
            default {
                Write-Host "`n"
                Write-Warning "Sorry, I didn't understand that!" 
                Write-Host "`n"                
            }
        }
    }
}

Function ImportContacts() {
    $FilePath = "$PSScriptRoot\forImport\contacts.csv"
    Write-Warning "Please ensure that $FilePath is filled out correctly before continuing!"
    Write-Host "The correct headings are Email,Name & Show"
    Write-Host "Show must be set as TRUE or FALSE, which defines whether or not to display it in the Global Address List"
    $proceed = $false
    $addtodist = $false

    $option = Read-Host "Are you ready to continue? [Y/N] (Selecting No will take you back to the main menu)"
    switch ($option) {
        Y { $proceed = $true }
        N { $proceed = $false }
        default { $proceed = $false }
    }

    if ($proceed) {

        $distlistoption = Read-Host "Do you want to add these contacts into a distribution group? [Y/N]"
        switch($distlistoption) {
            Y { $addtodist = $true }
            N { $addtodist = $false}
            default { $addtodist = $false }
        }

        if ($addtodist) {
            $distlist = Read-Host "Please enter the email address of the distribution list you would like to add these contacts to"
        }
        $csv = Import-CSV $FilePath
        ForEach ($contact in $csv)
        { 
            Write-Host "Processing" $contact.Name $contact.Email

            Write-Host "`n-`nAdding contact...`n"
            New-MailContact -Name $contact.Name -DisplayName $contact.Name -ExternalEmailAddress $contact.Email
            
            if ($contact.Show -eq "FALSE") {
                Write-Host "`n"
                Write-Host "Hiding from address list..."

                Set-MailContact $contact.Email -HiddenFromAddressListsEnabled $true
            }
            
            if ($distlist) {
                Write-Host "Adding to distribution list..."
                Add-DistributionGroupMember -Identity $distlist -Member $contact.Email
            }

            Write-Host "Done.`n-"
        }
        Write-Host "`n"
        Write-Host "Successfully imported contacts." -ForegroundColor Green
        Write-Host "`n"
    } else {
        Write-Host "Returning to main menu..."
    }
    
}

Function ImportToDistList() {
    $FilePath = "$PSScriptRoot\forImport\distListMembers.csv"
    Write-Warning "Please ensure that $FilePath is filled out correctly before continuing!"
    Write-Host "The correct headings are Email,DistList"

    $proceed = $false

    $option = Read-Host "Are you ready to continue? [Y/N] (Selecting No will take you back to the main menu)"
    switch ($option) {
        Y { $proceed = $true }
        N { $proceed = $false }
        default { $proceed = $false }
    }

    if ($proceed) {
        $csv = Import-CSV $FilePath
        ForEach ($user in $csv)
        {
            $email = $user.Email
            $distlist = $user.DistList
            Write-Host "Adding $email to $distlist..."
            Add-DistributionGroupMember -Identity $distList -Member $email
        }
        Write-Host "`n"
        Write-Host "Successfully imported users to the distribution list(s)." -ForegroundColor Green
        Write-Host "`n"
    } else {
        Write-Host "Returning to main menu..." 
    }
}

Function ImportDelegates() {
    $FilePath = "$PSScriptRoot\forImport\delegatePermissions.csv"
    Write-Warning "Please ensure that $FilePath is filled out correctly before continuing!"
    Write-Host "The correct headings are Mailbox,User,FullAccess,SendAs"
    Write-Host "To assign Full Access permissions put an X in this column. To assign Send As permissions, put an X in this column"

    $proceed = $false

    $option = Read-Host "Are you ready to continue? [Y/N] (Selecting No will take you back to the main menu)"
    switch ($option) {
        Y { $proceed = $true }
        N { $proceed = $false }
        default { $proceed = $false }
    }

    if ($proceed) {
        $csv = Import-CSV $FilePath
        ForEach ($mailbox in $csv)
        {
            if ($mailbox.FullAccess -eq "x") {
                Write-Host "User $($mailbox.User) to be added Full Access permissions to $($mailbox.Mailbox)"
                # Below adds the user read/manage access
                Add-MailboxPermission -Identity $mailbox.Mailbox -User $mailbox.User -AccessRights FullAccess -InheritanceType All
            }
            # Check if Send As permission is required.
            if ($mailbox.SendAs -eq "x") {
                Write-Host "User $($mailbox.User) to be added Send As permissions to $($mailbox.Mailbox)"
                # Below adds the user Send-As access
                Add-RecipientPermission -Identity $mailbox.Mailbox -Trustee $mailbox.User -AccessRights SendAs -Confirm:$false
            }
        }
        Write-Host "`n"
        Write-Host "Successfully imported and added delegate(s) to the mailbox(es)" -ForegroundColor Green
        Write-Host "`n"
    } else {
        Write-Host "Returning to main menu..." 
    }
}

Function ExitScript() { # Exits the script and cleans up
    Write-Host "`n"
    Write-Host "Thank you for using Niall's Exchange script! Goodbye." -ForegroundColor Green
    Write-Host "`n"
    Disconnect-ExchangeOnline -Confirm:$false
}


clear-host
Write-Host "Welcome to Niall's Exchange script."
Write-Host "First we will need to connect to Exchange..."
$Exchange = $false

# Clear any stale sessions
Disconnect-ExchangeOnline -Confirm:$false
# Create new session
Connect-ExchangeOnline -ShowBanner:$false

$connectinfo = Get-ConnectionInformation
#Confirm connection to Exchange
if ($connectinfo.Name -like "ExchangeOnline*" -and $connectinfo.State -eq 'Connected') {
    $Exchange = $true
}

if ($Exchange) { # Run script
    Write-Host "Connected to Exchange!" -ForegroundColor Green
    Write-Host "`n"
    DisplayMenu
} else { # Exit script
    Write-Warning "Connecting to Exchange failed."
    ExitScript
}

# (c) 2023 Niall Davies