#Cenk_BOZ
#Administrations_Tools_Exchange2013

$menu=@"

1 Show info database 
2 Show info history mailbox
3 Show info mailbox 
4 Show info quota
5 Show info forwarding
___________________________
6 Test MAPI Connectivity
Q Quit

Select a task by number or Q to quit
"@

cls
     Write-Host "Administrations_Tools_Exchange2013" -ForegroundColor Cyan
     $alias = Read-Host -Prompt "Enter LDAP"
     $r = Read-Host $menu

Switch ($r) {

#Info Database
"1" {
    Clear-Host
    Write-Host "Getting info database" -ForegroundColor Green
    Get-Mailboxdatabase * | % { $_.name }| % {$lib = 0; Get-mailbox -Database "$_" -ResultSize Unlimited -warningAction SilentlyContinue | % {$lib ++}; $out = "$_.database : " + "$lib" ; $out}
	
}

#Info mailbox History
"2" {
    Clear-Host
    Write-Host "Getting forwarding information" -ForegroundColor Green
    Get-MailboxStatistics -Identity $alias -IncludeMoveHistory | Format-List
	
}

#Info mailbox
"3" {
    Clear-Host
    Write-Host "Getting mailbox information" -ForegroundColor Green
    Get-Mailbox -Identity $alias | Format-List EmailAddresses
	
}

#Info Qota
"4" {
    Clear-Host
    Write-Host "Getting quota information" -ForegroundColor Green
    Get-Mailbox -Identity $alias | Format-List IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults
	
}

#Info Forwarding
"5" {
    Clear-Host
    Write-Host "Getting forwarding information" -ForegroundColor Green
    Get-Mailbox -Identity $alias | Format-List ForwardingSMTPAddress,DeliverToMailboxandForward
	
}

#Test MAPI Connectivity
"6" {
    Clear-Host
    Write-Host "Getting Test MAPI Connectivity" -ForegroundColor Green
    Test-MAPIConnectivity -Identity $alias | Format-List
	
}
default {
    Clear-Host
    Write-Host "Quitter" -ForegroundColor Yellow
	
 }
} #end switch

