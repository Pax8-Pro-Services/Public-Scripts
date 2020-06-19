<#
  _____            ___  
 |  __ \          / _ \ 
 | |__) |_ ___  _| (_) |
 |  ___/ _` \ \/ /> _ < 
 | |  | (_| |>  <| (_) |
 |_|   \__,_/_/\_\\___/ 
                       
.SYNOPSIS
Collects data necessary for quoting your mailbox migration.

(c) 2019 Pax8

#>
#Requires -Version 5.0
#Requires -PSEdition Desktop
Write-Host "Welcome to the Pax8 Professional Services - Migration Discovery Tool!" -ForegroundColor Green;

$isExo = Read-Host "Is this discovery for an Office 365 to Office 365 Migration? (y/n)"
$pDomain = Read-Host "Enter the primary domain being migrated"
$isUsingNewEOModule = $false;
function Connect-EOCustom {
    if(Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
        #The new module is installed
        Write-Host "You have the modern management shell installed, it will be used to connect to EO!" -ForegroundColor Green;
        Import-Module ExchangeOnlineManagement;
        $isUsingNewEOModule = $true;
        Connect-ExchangeOnline;
    } else {
        #the new module is not installed
        Write-Host "NOTICE: You are not using the modern management shell. The old shell will soon be deprecated, but the script will continue for now." -ForegroundColor Yellow;
        Write-Host "It is strongly recommended that you install the new shell using Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow;
        Write-Host "[!!] Certain tests will be skipped as they are written for the new module!" -ForegroundColor Red;
        Pause;
        $UserCredential = Get-Credential;
        try {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
            Import-PSSession $Session;
        } catch {
            Write-Host "Couldn't connect to Exchange Online." -ForegroundColor red;
            Pause;
            exit;
        }
        
    }
}

if($isExo -eq "y" -or $isExo -eq "Y") {
    $eo = $true;
} else {
    $eo = $false;
}

if($eo) {
    Write-Host "NOTE: This discovery tool only collects data from Exchange Online. OneDrive, SharePoint, and Teams require furuther discovery. Please let your Wingman know if you need information on migrating that information." -ForegroundColor Yellow;
    pause;
}

if($eo) {

    Connect-EOCustom;
    
} else {
    if(Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
        Write-Host "Verified that you're using Exchange Management Shell" -ForegroundColor Green;
    } else {
        Write-Host "[!] This tool must be ran from the Exchange Management Shell." -ForegroundColor Red;
        Pause;
        exit;
    }
}


$desktopPath = [Environment]::GetFolderPath("Desktop");

$dataFolder = "$desktopPath\$pDomain - MigrationData";

if((Test-Path -Path $dataFolder) -eq $false) {
    mkdir $dataFolder;
}

Write-Host "Created the following folder for this discovery: $dataFolder" -ForegroundColor Green;
Start-Sleep -Seconds 2;
class MailboxCsvRow {
    [object] ${DisplayName}
    [object] ${FirstName}
    [object] ${LastName}
    [object] ${PrimarySmtpAddress}
    [object] ${Size}
    [object] ${RecipientTypeDetails}
}

class BTRow {
    [object] ${FirstName}
    [object] ${LastName}
    [object] ${EmailAddress}
    [object] ${UserPrincipalName}
}

Write-Host "Doing initial data queries. This may take a moment, please be patient." -ForegroundColor Green;
Write-Progress -Activity "Mailbox Data Gathering" -Status "Getting Mailboxes";

$mailboxes = Get-Recipient -ResultSize Unlimited;
Write-Progress -Activity "Mailbox Data Gathering" -Status "Getting Distribution Lists";
$distributionLists = Get-DistributionGroup -ResultSize Unlimited | Select-Object Name,DisplayName,Alias,PrimarySmtpAddress,ManagedBy;
Write-Host "Exporting distribution lists to $dataFolder\distiLists.csv" -ForegroundColor Cyan;
$distributionLists | Export-Csv -Path "$dataFolder\distiLists.csv" -NoTypeInformation;
Write-Host "Handling mailboxes" -ForegroundColor Cyan;

$outputMailboxArray = @();
$outputBittitanArray = @();

Write-Progress -Activity "Mailbox Data Gathering" -Status "Building Reports";

foreach ($mailbox in $mailboxes) {
    $em = $mailbox.PrimarySmtpAddress;
    $firstName = $mailbox.FirstName;
    $lastName = $mailbox.LastName;
    if($eo) {
        if($mailbox.RecipientTypeDetails -eq "UserMailbox" -or $mailbox.RecipientTypeDetails -eq "GroupMailbox" -or $mailbox.RecipientTypeDetails -eq "SharedMailbox") {
            if($isUsingNewEOModule) {
                $mailboxId = $mailbox.ExternalDirectoryObjectId
                $mbSize = (Get-EXOMailboxStatistics $mailboxId).TotalItemSize.Value;
            } else {
                $mbSize = "Not Measured Due to Legacy PowerShell"
            }
            
        } else {
            $mbSize = "Not Measured due to Mailbox Type"
        }
        
    } else {
        $mbSize = "Not Measured"
    }
    

    $bittitanRow = [BTRow]::new();
    $bittitanRow.FirstName = $firstName;
    $bittitanRow.LastName = $lastName;
    $bittitanRow.EmailAddress = $mailbox.PrimarySmtpAddress;
    $bittitanRow.UserPrincipalName = $mailbox.PrimarySmtpAddress;
    $outputBittitanArray += $bittitanRow;
    
    $mailboxRow = [MailboxCsvRow]::new();
    $mailboxRow.DisplayName = $mailbox.DisplayName;
    $mailboxRow.FirstName = $firstName;
    $mailboxRow.LastName = $lastName;
    $mailboxRow.PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
    $mailboxRow.Size = $mbSize;
    $mailboxRow.RecipientTypeDetails = $mailbox.RecipientTypeDetails;
    $outputMailboxArray += $mailboxRow;
    Write-Host "Did $em" -ForegroundColor Gray;

}
Write-Host "Getting distribution list membership breakdown." -ForegroundColor Green;
Write-Progress -Activity "Mailbox Data Gathering" -Status "Getting DL Memberships";

$dlMembersArray=@()
$groups = Get-DistributionGroup -ResultSize Unlimited
$totalgroups = $groups.Count
$i = 1
$groups | ForEach-Object {
    Write-Progress -activity "Processing $_.DisplayName" -status "$i out of $totalgroups completed"
    $group = $_
    Get-DistributionGroupMember -Identity $group.Name -ResultSize Unlimited | ForEach-Object {
    $member = $_
    $dlMembersArray += New-Object PSObject -property @{
        GroupName = $group.DisplayName
        Member = $member.Name
        EmailAddress = $member.PrimarySMTPAddress
        RecipientType= $member.RecipientType
        }
}
$i++
}
Write-Progress -Activity "Mailbox Data Gathering" -Status "Saving Reports";
Write-Host "Exporting DL members to $dataFolder\dlMembers.csv" -ForegroundColor Yellow;
$dlMembersArray | Export-CSV "$dataFolder\dlMembers.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Exporting mailbox data to $dataFolder\mailboxes.csv" -ForegroundColor Cyan;
$outputMailboxArray | Export-Csv -Path "$dataFolder\mailboxes.csv" -NoTypeInformation;
$outputBittitanArray | Export-Csv -Path "$dataFolder\bt.csv" -NoTypeInformation;

Write-Host "Done with base mailbox data." -ForegroundColor Green;
Write-Host "============================================" -ForegroundColor Cyan;
Write-Host "The next step is public folder data gathering. Please let us know if you'd like Pax8 to perform a public folder migration." -ForegroundColor Yellow;
Write-Progress -Activity "Mailbox Data Gathering" -Status "Done" -Completed;
$doPublicFolderGathering = Read-Host "Include public folders (y/n)?";

if($doPublicFolderGathering -eq "y" -or $doPublicFolderGathering -eq "Y") {
    Write-Progress -Activity "Public Folder Data Gathering" -Status "Getting Public Folders";
    Write-Host "Getting public folder data. This may take a moment." -ForegroundColor Green;
    $pf = Get-PublicFolderStatistics -ResultSize Unlimited | Select-Object Name, ItemCount, TotalItemSize, LastUserAccessTime, LastUserModificationTime;
    $pfs = Get-PublicFolder -Recurse | Select-Object Identity;
    Write-Progress -Activity "Public Folder Data Gathering" -Status "Writing Reports";
    Write-Host "Exporting public folder data to $dataFolder\pf.csv and $dataFolder\pf-structure.csv";
    $pf | Export-Csv -Path "$dataFolder\pf.csv" -NoTypeInformation;
    $pfs | Export-Csv -Path "$dataFolder\pf-structure.csv" -NoTypeInformation;
    Write-Progress -Activity "Public Folder Data Gathering" -Status "Getting Public Folders" -Completed;
} else {
    Write-Host "Skipping public folders" -ForegroundColor Yellow;
}
$totalMailboxes = ($outputMailboxArray | Measure-Object).Count;

Write-Host "We're all set, $totalMailboxes mailboxes discovered!" -ForegroundColor Green;
if($doPublicFolderGathering -eq "y" -or $doPublicFolderGathering -eq "Y") {
    Write-Host "Public folder data has been collected, but must be processed manually for accurate quoting." -ForegroundColor Yellow;
}
pause;