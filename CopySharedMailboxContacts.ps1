#
#   CopySharedMailboxContacts.ps1
#
#   Script to create Outlook contacts from AD User objects
#   based on group memberships
#
#   8/07/2024: V1.0 Initial Release
#
#   Script Author: Christian Schindler, NTx BackOffice Consulting Group Gmbh
#
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [System.IO.FileInfo]
    $ConfigFile
)
$Config = Get-Content -Path $ConfigFile | ConvertFrom-Json

# Variable definition
[string]$ContactSourceMailbox = $Config.ContactSourceMailbox
[string]$GroupForContactDestination = $Config.GroupForContactDestination
[string]$ContactFolderName = $config.ContactFolderName
[string]$BasePath = "C:\Admin\Scripts"
$Script:NoLogging
$ExchangeNameSpace = "winmail.e-control.loc"
[string]$LogfileFullPath = Join-Path -Path $BasePath (Join-Path $ContactFolderName ("CopySharedMailboxContacts_" + $($ContactSourceMailbox.Split("@")[0]) + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now))

# Disable the Active Directory Provider
$Env:ADPS_LoadDefaultDrive = 0

# Module loading
Import-Module -Name ActiveDirectory

function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured LogfileFullPath variable to identify the logfile and NoLogging to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$Mailbox,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )
    # Prefix the string to write with the current Date and Time, add error message if present...

    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2} Error: {3}" -f [DateTime]::Now, $Mailbox, $Message, $ErrorInfo.Exception.Message
    }

    elseif ($Mailbox)
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2}" -f [DateTime]::Now, $Mailbox, $Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}" -f [DateTime]::Now, $Message
    }
    if (!$Script:NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogfileFullPath -PathType Leaf))
        {
            New-Item -ItemType File -Path $Script:LogfileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
            Add-Content -Value "Logging started." -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    else
    {
        Write-Host $logLine
    }
}

function Load-EWSManagedAPI
{
    ## Load Managed API dll
    ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
    try
    {
        $EWSDLL = (($(Get-ItemProperty -ErrorAction Stop -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
        Write-LogFile -Message "EWS Managed API DLL found"
    }
    catch
    {
        Write-LogFile -Message "EWS Managed API DLL not found." -ErrorInfo $_
        exit
    }

    try
    {
        Import-Module $EWSDLL -ErrorAction Stop
        Write-LogFile -Message "EWS Managed API successfully loaded."
    }
    catch
    {
        Write-LogFile -Message "EWS Managed API could not be loaded." -ErrorInfo $_
        exit
    }
}

function Connect-Exchange
{
    #
    # Function to connect to a mailbox via EWS impersonation
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]$MailboxName
    )

    ## Set Exchange Version
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

    ## Create Exchange Service Object
    $exservice = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

    ## Use the Default (logged On) credentials
    $exservice.UseDefaultCredentials = $true
    #$exservice.Credentials = New-Object Net.NetworkCredential($username, $password)

    # Set EWS URL
    [system.URI]$uri = "https://" + $ExchangeNameSpace + "/ews/exchange.asmx"
    $exservice.Url = $uri

    ## Optional section for Exchange Impersonation
    $exservice.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)

    return $exservice
}

function GetSourceContacts
{
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)] [string]$MailboxName
    )

    $Connection = Connect-Exchange -MailboxName $MailboxName

    #$SourceContacts = New-Object System.Collections.ArrayList
    $SourceContacts = @()

    # Connect to the mailbox
    $ContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Connection, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

    $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Contact")
	
    #Define ItemView to retrive just 1000 Items    
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
	$fiItems = $null

    $fiItems = $Connection.FindItems($ContactsFolder.Id,$SfSearchFilter,$ivItemView)

    if($fiItems.Items.Count -gt 0)
    {
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        [Void]$Connection.LoadPropertiesForItems($fiItems,$psPropset)

        foreach($Item in $fiItems.Items){      
            $SourceContacts += $Item
        }
    }

    If ($SourceContacts.Count -gt 0)
    {
        Write-LogFile -Message "Successfully retrieved $($SourceContacts.Count) contacts from Sourcemailbox $($MailboxName)"
        Return $SourceContacts
    }

    else
    {
        Write-LogFile -Message "No contacts found in source mailbox. Exiting..."
        Exit
    }
}

function GetContactDestination
{
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$GroupForContactDestination
    )

    #
    # Retrieve group members of specified group and store them in an array
    #

    $Members = @()

    try
    {
        $Members = Get-ADGroupMember -Identity $GroupForContactDestination -ErrorAction Stop
        Write-LogFile -Message "Successfully retrieved $($members.Count) destination mailboxes for contact sync from group $($GroupForContactDestination)"
        Write-LogFile -Message "Listing members retrieved:"
    }

    catch
    {
        Write-LogFile "Function GetContactDestination: Unable to retrieve members from group $($GroupForContactDestination)." -ErrorInfo $_
        throw $_
    }

    #
    # Retrieve required properties of group members and store them in an array
    #


    $DestinationMailboxes = @()

    foreach ($member in $Members)
    {
        $user = Get-ADUser -Identity $member.SamAccountName -Properties mail, displayname
        Write-LogFile -Message "Retrieved destination mailbox $($user.mail)"
        $DestinationMailboxes += $user
    }

    # Copy members to an arraylist so we can modify it in the loop

    return $DestinationMailboxes

    #
    # If the attribute "extensionAttribute1" contains "nocontact" remove the user from the arraylist
    #

    #foreach ($user in $DestinationMailboxes)
    #{
    #    if ($user.extensionAttribute1 -eq "nocontact")
    #    {
    #        $Finalmembers.Remove($user)
    #    }
    #}
}
function FolderExists 
{
    param (
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Connection,
        [Microsoft.Exchange.WebServices.Data.Folder]$ContactsFolder,
        [string]$ContactFolderName
    )

    # Define a Search folder that is going to do a search based on the DisplayName of the folder
    $SfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $ContactFolderName)

    # Define a folder view
    $Folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)

    # Do the Search
    $findFolderResults = $Connection.FindFolders($ContactsFolder.Id, $SfSearchFilter, $Folderview)

    return $findFolderResults.TotalCount
}
function ManageContactFolder
{
    [CmdLetBinding()]

    param(
        [Parameter(Mandatory = $true)] [string]$MailboxName,
        [Parameter(Mandatory = $false)] [String]$ContactFolderName,
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Connection
    )

    $FolderClass = "IPF.Contact"

    # Connect to the mailbox
    #$Connection = Connect-Exchange -MailboxName $MailboxName

    # Bind to the MsgFolderRoot folder
    #$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName)
    #$EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Connection, $folderid)

    # Bind tot the contacts folder
    $ContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Connection, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

    #Define Folder Veiw Really only want to return one object
    $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)

    # Define the new folder an it's properties
    $NewFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($Connection)
    $NewFolder.DisplayName = $ContactFolderName
    $NewFolder.FolderClass = $FolderClass
    $EWSParentFolder = $null

    # Define a Search folder that is going to do a search based on the DisplayName of the folder
    $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $ContactFolderName)

    # Do the Search
    $findFolderResults = $Connection.FindFolders($ContactsFolder.Id, $SfSearchFilter, $fvFolderView)

    # If the search was not succesful
    if ($findFolderResults.TotalCount -eq 0)
    {
        Write-LogFile -Mailbox $MailboxName -Message "Folder Doesn't Exist"

        # Try creating the folder as a subfolder of the "Contacts" folder
        try
        {
            $NewFolder.Save($ContactsFolder.Id)
            Write-LogFile -Mailbox $MailboxName -Message "Folder $ContactFolderName successfully created."
        }
        catch
        {
            Write-LogFile -Mailbox $MailboxName -Message "Could not create folder $ContactFolderName." -ErrorInfo $_
            Throw $_
        }
    }

    # if the search was successful
    else
    {
        Write-LogFile -Mailbox $MailboxName -Message "Folder $ContactFolderName already exists."
        Write-LogFile -Mailbox $MailboxName -Message "Deleting folder and recreating."

        # Try deleting the folder
        try
        {
            $findFolderResults.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
            Write-LogFile -Mailbox $MailboxName -Message "Folder $ContactFolderName successfully deleted."

            do { $findFolderResults1 = $Connection.FindFolders($ContactsFolder.Id, $SfSearchFilter, $fvFolderView); Start-Sleep -Seconds 4 }
            until ($findFolderResults1.TotalCount -eq 0)
            $folderdeleted = $true
        }

        catch
        {
            Write-LogFile -Mailbox $MailboxName -Message "Could not delete folder $ContactFolderName." -ErrorInfo $_
            Throw $_
        }

        # If the existing folder was successfully deleted
        if ($folderdeleted)
        {
            Try
            {
                $NewFolder.Save($ContactsFolder.Id)
                Write-LogFile -Mailbox $MailboxName -Message "Folder $ContactFolderName successfully created."
            }
            catch
            {
                Write-LogFile -Mailbox $MailboxName -Message "Could not create folder $ContactFolderName." -ErrorInfo $_
                Throw $_
            }
        }

    }

    return $NewFolder

}

function CreateContact 
{
    [CmdLetBinding()]
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$folder,
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Connection,
        [string]$GivenName,
        [String]$Surname,
        [String]$DisplayName,
        [string]$Department,
        [string]$Office,
        [string]$telephoneNumber,
        [string]$Mobile,
        [string]$mail,
        [string]$title,
        [byte[]]$Thumbnailphoto
    )

    # Check if the contact already exists in the folder
    $contactExists = $folder.FindItems("DisplayName -eq '$DisplayName'", 1).TotalCount -ne 0

    if (-not $contactExists) {
        # Create the contact object in the current mailbox
        $Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $Connection

        # Set contact properties
        $Contact.GivenName = $GivenName
        $Contact.Surname = $Surname
        $Contact.Subject = $DisplayName
        $Contact.FileAs = $DisplayName
        $Contact.DisplayName = $DisplayName
        $Contact.Department = $Department
        $Contact.OfficeLocation = $Office
        $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $telephoneNumber
        $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $Mobile
        $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $mail
        $Contact.JobTitle = $title

        # If a photo exists, store the photo in the contact
        if ($Thumbnailphoto)
        {
            $Contact.SetContactPicture($Thumbnailphoto)
        }

        # Save the new contact object to the ECA-Contact Folder
        try
        {
            $Contact.Save($folder.Id)
            Write-LogFile -Mailbox $($connection.ImpersonatedUserId.Id) -Message "Successfully created contact $($Contact.Displayname)."
        }
        
        catch
        {
            Write-LogFile -Mailbox $($connection.ImpersonatedUserId.Id) -Message "Could not create Contact $($Contact.Displayname)" -ErrorInfo $_
            #Throw $_
        }
    }

    else {
        Write-LogFile -Mailbox $($connection.ImpersonatedUserId.Id) -Message "Contact $DisplayName already exists in the folder."
    }
}

#
# Main Script
#

# Load EWS Managed API
Load-EWSManagedAPI

# Retrieve Contacts from source
$SourceContacts = GetSourceContacts -MailboxName $ContactSourceMailbox

# Retrtieve mailboxes to store contacts in
$Mailboxes = GetContactDestination -GroupForContactDestination $GroupForContactDestination

if ($Mailboxes.Count -gt 0)
{
    Write-LogFile -Message "Looping through destination mailboxes..."
    Write-LogFile -Message "---------------------------------------------------------"
}

else
{
    Write-LogFile -Message "No destination mailboxes found. Exiting the script..."
    Exit
}

# Loop through list if mailboxes
foreach ($Mailbox in $Mailboxes)
{
    # Connect to Mailbox via EWS
    try
    {
        $Connection = Connect-Exchange -MailboxName $Mailbox.mail

        Write-LogFile -Mailbox $Mailbox.mail -Message "Successfully connected to mailbox"
    }
    catch
    {
        Write-LogFile -Mailbox $Mailbox.mail -Message "Unable to connect to mailbox" -ErrorInfo $_
    }

    # Delete and recreate Contacts folder
    $folder = ManageContactFolder -MailboxName $Mailbox.mail -ContactFolderName $ContactFolderName -Connection $Connection

    Write-LogFile -Message "Creating $($SourceContacts.Count) contacts in mailbox $($Mailbox.mail)"

    # Loop through contacts
    foreach ($contact in $SourceContacts)
    {
        # For each entry, create a new contact
        CreateContact -folder $Folder -Connection $Connection -GivenName $Contact.GivenName -Surname $Contact.Surname -DisplayName $contact.DisplayName -Department $Contact.Department -Office $Contact.physicalDeliveryOfficeName -telephoneNumber $Contact.telephoneNumber -Mobile $Contact.mobile -mail $Contact.EmailAddresses[0].Address -title $Contact.Title -Thumbnailphoto $Contact.thumbnailPhoto
    }

    Write-LogFile -Message "Finished creating contacts in mailbox $($Mailbox.Displayname)."
    Write-LogFile -Message "---------------------------------------------------------"

    # Cleanup
    $connection = $null
    $folder = $null
    $contact = $null
}