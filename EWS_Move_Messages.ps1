############################################################################################
#                                                                                          #
# The sample scripts are not supported under any Microsoft standard support                #
# program or service. The sample scripts are provided AS IS without warranty               #
# of any kind. Microsoft further disclaims all implied warranties including, without       #
# limitation, any implied warranties of merchantability or of fitness for a particular     #
# purpose. The entire risk arising out of the use or performance of the sample scripts     #
# and documentation remains with you. In no event shall Microsoft, its authors, or         #
# anyone else involved in the creation, production, or delivery of the scripts be liable   #
# for any damages whatsoever (including, without limitation, damages for loss of business  #
# profits, business interruption, loss of business information, or other pecuniary loss)   #
# arising out of the use of or inability to use the sample scripts or documentation,       #
# even if Microsoft has been advised of the possibility of such damages                    #
#                                                                                          #
# Author: doqi@microsoft.com                                                               #
############################################################################################
<#
     .SYNOPSIS
        The PowerShell script which can be used to move the messages (under the specific conditions for a mailbox 
        (Exchange online or Exchange onprem) to a specified folder.Support to run the script in PowerShell 5 only. 
    .DESCRIPTION
        The PowerShell script which can be used to move the messages (under the specific conditions for a mailbox 
        (Exchange online or Exchange onprem) to a specified folder.Support to run the script in PowerShell 5 only. 

    .Author
        Qi Dong (doqi@microsoft.com)
    .PARAMETER  
        Mailbox: the user mailbox to move the messages. Mandatory
        TargetFolderName: the target folder to move the messages to.Mandatory
        IsTargetFolderInArchiveMailbox: the target folder is in Archive or primary mailbox. Mandatory
        StartDateTime: the start datetime for the DateTimeReceived of the messages. It is your local datetime. Mandatory
        EndDateTime: the end datetime for the DateTimeReceived of the messages. It is your local datetime. Mandatory
        SubjectString: the substring of the subject to search the messages. Optional
        Sender: the sender email address of the subject to search. Optional
        MessageId: the message-Id of the email to search. Optional
        MaxItemsPerProcessing: the max number of messages is found by the search conditions (exclude the sender). Optional. The default is 5000
        WellKnowFolder: the targetted folder to search the message. Optional. The default is Inbox.https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.wellknownfoldername?view=exchange-ews-api    
        PageSize: use the EWS paging to control how many items are returned. Optional. The default is 100
        DeleteMode: the delete mode. Optional. The default is MoveToDeletedItems. https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.deletemode?view=exchange-ews-api
        IsConfirm: before the message delete, whetheer ask the user's confirmation. Optional. The default is false.
        ExchangeOnline: whether it is an EXO mailbox. Optional. The default is false.
    .EXAMPLE
        $mailbox = 'test@freeguys13.onmicrosoft.com';
        $TargetFolderName = 'TestFolder';
        $IsTargetFolderInArchiveMailbox = $true;
        #Find an email message to remove with the following search conditions
        $startDateTime = "2021-02-20";
        $endDateTime = "2022-04-24";
        .\MoveMessages.ps1 -Mailbox $mailbox -TargetFolderName $TargetFolderName -IsTargetFolderInArchiveMailbox $IsTargetFolderInArchiveMailbox -StartDateTime $startDateTime -EndDateTime $endDateTime -IsConfirm -ExchangeOnline;

    .EXAMPLE
        Use the Windows Integration authentication and the service account has the impersonation permission to access another user mailbox
        https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-configure-impersonation
        $mailbox = 'RiquelTest@huizhan.msftonlinelab.com';
        $TargetFolderName = 'TestFolder';
        $IsTargetFolderInArchiveMailbox = $true;
        #Find an email message to remove with the following search conditions
        $startDateTime = "2021-02-20";
        $endDateTime = "2022-04-24";
        $SubjectString = "test";
        $Sender = 'test@contoso.com';
        .\MoveMessages.ps1 -Mailbox $mailbox -StartDateTime $startDateTime -EndDateTime $endDateTime -SubjectString $SubjectString -Sender $Sender -IsConfirm;        
#>

#Todo: need to supprt the EWS trace feature and the ItemTraversal Enum in FindItems call. Also need to add the parameter validation for DeleteMode, WellKnowFolder

# Import the required libraries. Require PowerShell 5 version
using module ".\Microsoft.Exchange.WebServices.dll";
using module '.\Microsoft.IdentityModel.Clients.ActiveDirectory.dll';

[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $True)]
    [String] $Mailbox,
    [Parameter(Position = 1, Mandatory = $True)]
    [String] $TargetFolderName,
    [Parameter(Position = 2, Mandatory = $True)]
    [bool] $IsTargetFolderInArchiveMailbox,
    [Parameter(Position = 3, Mandatory = $True)]
    [DateTime] $StartDateTime,
    [Parameter(Position = 4, Mandatory = $True)]
    [DateTime] $EndDateTime,     
    [Parameter(Position = 5, Mandatory = $False)]
    [String] $SubjectString,
    [Parameter(Position = 6, Mandatory = $False)]
    [String] $Sender = "",
    [Parameter(Position = 7, Mandatory = $False)]
    [String] $MessageId = "",
    [Parameter(Position = 8, Mandatory = $False)]
    [int]$MaxItemsPerProcessing = 5000,
    [Parameter(Position = 9, Mandatory = $False)] 
    [Int]$WellKnowFolder = 4,
    [Parameter(Position = 10, Mandatory = $False)]
    [int]$PageSize = 100,
    [Parameter(Position = 11, Mandatory = $False)] 
    [int]$DeleteMode = 2,
    [Parameter(Position = 12, Mandatory = $False)]
    [Switch]$IsConfirm = $False,
    [Parameter(Position = 13, Mandatory = $False)]
    [Switch]$ExchangeOnline = $False
)

function ShowErrorDetails {
    param(
        $ErrorRecord = $Error[0]
    )
    $ErrorRecord | Format-List -Property * -Force
    $ErrorRecord.InvocationInfo | Format-List -Property *
    $Exception = $ErrorRecord.Exception
    for ($depth = 0; $null -ne $Exception; $depth++) {
        "$depth" * 80                                               
        $Exception | Format-List -Property * -Force                 
        $Exception = $Exception.InnerException                      
    }
}

Function GetExchangeService {
    Param (
        [Parameter(Position = 0, Mandatory = $True)]
        [String] $mailboxToProcess
    )
    [Microsoft.Exchange.WebServices.Data.ExchangeService]$service = $null;
    try {

        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1;
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion);

        If ($ExchangeOnline) {
            #Refer to client credential auth (https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
            $tenantID = "7e0ce6d7-b890-47e0-9812-a1afd691bf5f";
            $authString = "https://login.microsoftonline.com/$tenantID";
            $appId = "4b21b615-dbdb-4fee-ae48-15eff6726662";
            $appSecret = "";
            $uri = [system.URI] "https://outlook.office365.com/EWS/Exchange.asmx";
            $creds = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $appId, $appSecret
            $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext"-ArgumentList $authString
            $context = $authContext.AcquireTokenAsync("https://outlook.office365.com/", $creds).Result;

            $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $context.AccessToken;
            $service.url = $uri;
            $service.UserAgent = "MyEWSClientAgent-DeletingEmails";
        } 
        Else {
            $service.Credentials = [System.Net.CredentialCache]::DefaultCredentials;
            $service.AutodiscoverUrl($mailboxToProcess);
        }
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress, $mailboxToProcess);
        $service.HttpHeaders.Add("X-AnchorMailbox", $mailboxToProcess);
    }
    catch {
        ShowErrorDetails;
        $service = $null;
    }
    $service;
    return;
}

Function GetSearchFilter {
    param (
        [Parameter(Position = 0, Mandatory = $True)]
        [DateTime] $StartDateTime,
        [Parameter(Position = 1, Mandatory = $True)]
        [DateTime] $EndDateTime,        
        [Parameter(Position = 2, Mandatory = $False)]
        [String] $SubjectString,
        [Parameter(Position = 3, Mandatory = $False)]
        [String] $MessageId
    )
    $searchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
    $sf1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $startDateTime);
    $sf2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $endDateTime);
    $searchFilterCollection.Add($sf1);
    $searchFilterCollection.Add($sf2);

    if (-Not [String]::IsNullOrEmpty($SubjectString)) {
        $sf3 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $SubjectString, [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase);
        $searchFilterCollection.Add($sf3);
    }

    <# this way doesn't work, so will filter the emails by sender
    if(-Not [String]::IsNullOrEmpty($Sender)){
        $fromAddress = New-Object Microsoft.Exchange.WebServices.Data.EmailAddress -ArgumentList $Sender;
        $sf4 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender,$fromAddress); 
        $searchFilterCollection.Add($sf4);
    }
    #>

    if (-Not [String]::IsNullOrEmpty($MessageId)) {
        $MessageID = "<$($MessageID.Trim('<','>'))>";
        $sf5 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId, $MessageId) 
        $searchFilterCollection.Add($sf5);
    }
    $Global:searchFilter = $searchFilterCollection;
    return;
}

Function FindSourceFolder {
    param(
        [Parameter(Position = 0, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$exchangeWebService,
        [Parameter(Position = 1, Mandatory = $True)]
        [Int]$targetFolder  
    )
    $wellKnowFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$targetFolder;
    $sourceFolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId($wellKnowFolderName, $Mailbox);
    [Microsoft.Exchange.WebServices.Data.Folder]$sourceFolder = $null;
    try {
        $sourceFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeWebService, $sourceFolderID);
    }
    catch {
        ShowErrorDetails;
        $sourceFolder = $null;
    }
    $sourceFolder;
    return;
}

Function FindTargetFolder {
    param(
        [Parameter(Position = 0, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$exchangeWebService,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$targetFolder,
        [Parameter(Position = 2, Mandatory = $True)]
        [bool]$IsTargetFolderInArchiveMailbox
    )

    $FolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(2000);
    $FolderView.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::Id, [Microsoft.Exchange.WebServices.Data.FolderSchema]::TotalCount);
    $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
    if ($IsTargetFolderInArchiveMailbox) {
        $folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $FolderView);
    }
    else {
        $folders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $FolderView);
    }
    $targetMailboxFolder = $null;
    $folders | ForEach-Object {
        if ($PSItem.DisplayName -eq $targetFolder) {
            $targetMailboxFolder = $PSItem;
        }
    }
    $targetMailboxFolder;
    return;
}

Function FindMessages {
    param (
        [Parameter(Position = 0, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.Folder]$targetFolder,
        [Parameter(Position = 1, Mandatory = $True)]
        [int]$pageSize,
        [Parameter(Position = 2, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$searchFilter,
        [Parameter(Position = 3, Mandatory = $True)] # In each running, the maximum items are processed
        [int]$maxItemsPerProcessing,
        [Parameter(Position = 4, Mandatory = $false)]
        [String] $sender = "",
        [Parameter(Position = 1, Mandatory = $false)]
        [Object]$objList
    )   
    $offset = 0;
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(($pageSize + 1), $offset);
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, 
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender);
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow;
    
    # refer to paging implementation https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-paged-searches-by-using-ews-in-exchange
    [bool]$moreItems = $true;
    [Microsoft.Exchange.WebServices.Data.ItemId]$anchorId = $null;
    $itemCount = 0;

    $idsList = $objList -as 'System.Collections.Generic.List[System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]]';
    if ($null -eq $idsList) {
        Write-Host "the Ids List is null. Exit";
        exit;
    }

    while ($moreItems) {
        $result = $targetFolder.FindItems($searchFilter, $view);
        $moreItems = $result.MoreAvailable;

        if (($null -eq $result) -or ($null -eq $result.Items) -or ($result.Items.Count -eq 0)) {
            break;    
        }
    
        if ($moreItems -and ($null -ne $anchorId)) {
            $testId = $result.Items[0].Id;
            if ($testId -ne $anchorId) {
                Write-Host "The collection has changed while paging. Some results may be missed.";
            }
        }
        
        if ($moreItems) {
            $view.Offset = $view.Offset + $pageSize;
        }
        $anchorId = $result.Items[$result.Items.Count - 1].Id;
    
        $itemCount = $itemCount + $result.Items.Count;
        if ($itemCount -ge $maxItemsPerProcessing) {
            $moreItems = $false;
        }

        $ids = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]';
        $result.Items | ForEach-Object {
            $mail = $PSItem -as [Microsoft.Exchange.WebServices.Data.EmailMessage];
            if ($null -ne $mail) {
                if (([String]::IsNullOrEmpty($sender)) -or ($mail.Sender.Address -eq $sender)) {
                    $mail | Select-Object Subject, InternetMessageId, Sender, DateTimeReceived | Out-Host;
                    $ids.Add($PSItem.Id);                    
                }
            }
        }
    
        if ($ids.Count -gt 0) {
            $idsList.Add($ids);       
        }    
        
        $ids = $null;
        $result = $null;
    }
}

Function MoveMessages {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param (
        [Parameter(Position = 0, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$exchangeWebService,
        [Parameter(Position = 1, Mandatory = $True)]
        [Object]$objList,
        [Parameter(Position = 2, Mandatory = $True)]
        [Microsoft.Exchange.WebServices.Data.Folder]$targetFolder
    )
    if ($PSCmdlet.ShouldProcess("some found messages", "MoveMessages")) {
        $idsList = $objList -as 'System.Collections.Generic.List[System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]]';
        if ($null -eq $idsList) {
            Write-Host "the Ids List is null. Exit";
            exit;
        }
        if ($idsList.Count -gt 0) {
            $idsList | ForEach-Object {
                [System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]]$ids = $PSItem;
                if ($ids.Count -gt 0) {
                    $exchangeWebService.MoveItems($ids, $targetFolder.Id) | Out-Null;
                }
                [int]$sleepSeconds = Get-Random -Minimum 0 -Maximum 3;
                Start-Sleep -Seconds $sleepSeconds;
            }
        }
    }
}

#Start the execution
Set-StrictMode -Version 2;
$InformationPreference = "Continue";
$Global:ErrorActionPreference = "Stop";

[Microsoft.Exchange.WebServices.Data.ExchangeService]$service = GetExchangeService -mailboxToProcess $Mailbox;
[Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$Global:searchFilter = $null;
GetSearchFilter -StartDateTime $StartDateTime -EndDateTime $EndDateTime -SubjectString $SubjectString -MessageId $MessageId;

if ($null -eq $service) {
    Write-Host "Can not create the Exchange service object. Exit";
    exit;
}

if ($null -eq $global:searchFilter) {
    Write-Host "Can not create the search filter object. Exit";
    exit;
}

$sourceFolder = FindSourceFolder -exchangeWebService $service -targetFolder $WellKnowFolder;
if ($null -eq $sourceFolder) {
    Write-Host "Can not find the source folder. Exit";
    exit;
}

$targetMailboxFolder = FindTargetFolder -exchangeWebService $service -targetFolder $TargetFolderName -IsTargetFolderInArchiveMailbox $IsTargetFolderInArchiveMailbox;
if ($null -eq $targetMailboxFolder) {
    Write-Host "Can not find the targetted folder. Exit";
    exit;
}

$idsList = New-Object 'System.Collections.Generic.List[System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]]';
FindMessages -targetFolder $sourceFolder -pageSize $PageSize -searchFilter $Global:searchFilter -maxItemsPerProcessing $MaxItemsPerProcessing -Sender $Sender -objList $idsList;

if ($idsList.Count -gt 0) {
    if ($IsConfirm) {
        MoveMessages -exchangeWebService $service -objList $idsList -targetFolder $targetMailboxFolder -Confirm;
    }
    else {
        MoveMessages -exchangeWebService $service -objList $idsList -targetFolder $targetMailboxFolder;
    }
}