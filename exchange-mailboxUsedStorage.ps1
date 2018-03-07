<#
    AUTHER======: Garvey Snow
    VERSION=====: 1.0
    DESCRIPTION=: Pull's mailbox statistics from exchange server
                  -switch [string[]]$type - takes the following values
                                              • UserMailbox
                                              • LegacyMailbox
                                              • SharedMailbox
                                              • ResourceMailbox
                                              • LinkedMailbox
                                              • Disconnected Mailbox
    DEPENDANCIES: Requies access to exchange CMLETS
                  Use Exchange-Connect URL: 
	BUILD ENV===: Powershell Version 5.0.10586.117
    LICENCE=====: GNU GENERAL PUBLIC LICENSE
	UPDATED=====: 01/03/2018 @ 10:49AM
#>

function exchange-mailboxUsedStorage()
{
    
    param(
            [string[]]$type,
            [string[]]$exportpath,
            [string[]]$limit
         )

$allmailboxs = get-mailbox -RecipientTypeDetails $type -ResultSize unlimited

#///////////////////////////////
#// Custom HASH Table 
#// To Hold mailbox objects
#//////////////////////////////
$custom_object_array = @()

#//////////////////////////////
#// Custom Date String[]
#//////////////////////////////
$rawdate = (Get-Date | select datetime).datetime
$formated_date = $rawdate.ToString() -replace ' ','-' -replace ',','' -replace ':','-'

$total_mailbox_count = $allmailboxs.Count

foreach($mbx in $allmailboxs)
{
    
    #/////////////
    #// Varibales
    #/////////////
    $mailbox_Name = $mbx.Name
    $current_position = [array]::IndexOf($allmailboxs,$mbx) <# Get the current position of the loop#>
    $current_object_database = $mbx.Database <# Get the current user mailbox #>
    
    #///////////////////////////////////////////////////////////////
    #// Percentage of currently loop position 00.00 number format
    #///////////////////////////////////////////////////////////////
    $percent_complete = [math]::Round(([array]::IndexOf($allmailboxs,$mbx) / $allmailboxs.count * 100),2)

    #///////////////////////////////////////////////////
    #// Write-progess to screen with object information
    #///////////////////////////////////////////////////
    Write-Progress -Activity "Checking Statistic's" -Status "Polling $current_position/$total_mailbox_count $percent_complete% Complete" -Id 1 -PercentComplete $percent_complete -CurrentOperation "$mailbox_Name on Database $current_object_database"

    #////////////////////////////////////
    #// Mailbox stats of current object
    #///////////////////////////////////
    $mbx_stats = $mbx | Get-MailboxStatistics
    
    #////////////////////////////////////////////
    #// Manipulate string return only [int] valye
    #////////////////////////////////////////////
    if($mbx_stats.TotalItemSize -ne $null -or $mbx_stats.TotalItemSize.length -ne 0 )
    {
        $box_string = $mbx_stats.TotalItemSize.ToString()
        $raw_byes = $box_string.Split(" ")[2].replace("(", '') -replace ',',''
    }
    else
    {
        $raw_byes = 'Null Object'
    }

    #//////////////////////////////////////////////////////
    #// CUSTOM OBJECT - Holds properties from 
    #//                     - Mailbox Object
    #//                     - Mailbox Statistics Object 
    #///////////////////////////////////////////////////////
    $custom_object_array += New-Object -TypeName PSObject -Property @{
                                                                        SamAccountName = $mbx.SamAccountName
                                                                        OrganizationalUnit = $mbx.OrganizationalUnit
                                                                        PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                                                                        MailboxType = $mbx.RecipientTypeDetails
                                                                        CreationDate =  $mbx.WhenCreated
                                                                        lastUpdated = $mbx.WhenChanged
                                                                        LastLogonTime = $mbx_stats.LastLogonTime
                                                                        DatabaseName = $mbx_stats.DatabaseName
                                                                        DatabaseServer = $mbx_stats.ServerName
                                                                        DataBaseSize = $raw_byes
                                                                        ItemCount = $mbx_stats.ItemCount 
                                                                        DeletedItemCount = $mbx_stats.DeletedItemCount                     
                                                                     }
}
#//////////////////////////////
#// Generate custom save path
#//////////////////////////////
if($exportpath -eq $null)
{
    $PATH_EX = $env:USERPROFILE.ToString()+ "\Documents\exhange_report_space_used-$type-$formated_date.csv"
    $PATH_EX -replace ",","-"
    $PATH_EX_STR = $PATH_EX.ToString()

}
else
{
    $PATH_EX = $exportpath + "\exhange_report_space_used-$type-$formated_date.csv"
    $PATH_EX_STR = $PATH_EX.ToString()
}
   

    $custom_object_array | Sort-Object -Property DataBaseSize | Export-Csv -path $PATH_EX_STR -NoTypeInformation

    # NULL OBJECTS AND STRINGS
    $custom_object_array = $null
    $allmailboxs= $null
    $mbx_stats = $null


    write-host -ForegroundColor Green "EXPORTED .CSV to >> " -NoNewline; write-host -ForegroundColor Cyan $PATH_EX_STR

}



## USAGE
## -----
exchange-mailboxUsedStorage -type sharedmailbox,usermailbox -exportpath

