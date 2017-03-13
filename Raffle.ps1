<#
.Synopsis
   Raffle Time!
.DESCRIPTION
   Takes a csv file consisting of Name,Number of Entries and picks 1st, 2nd, & 3rd place winners
.EXAMPLE
   .\Raffle.ps1 C:\temp\raffle.csv
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$InFile

        
    )


$Entries = Import-Csv $InFile
$Tickets = @{}
$TicketNo = 0

#Fill in the Tickets
foreach ($Row in $Entries){
    
    for ($i = 1; $i -lt $row.Tickets; $i++){
        $Tickets.$TicketNo = $Row.Name
        $TicketNo++
        }    
}
$First = Get-Random -Maximum $TicketNo
$Second = Get-Random -Maximum $TicketNo
while ($Second -eq $First)
{
    $Second = Get-Random -Maximum $TicketNo
}
$Third = Get-Random -Maximum $TicketNo
while ($Third -eq $First -or $Third -eq $Second)
{
    $Third = Get-Random -Maximum $TicketNo
}

Write-Host "First place is $($Tickets[$First]) with Ticket number $($First)"
Write-Host "Second place is $($Tickets[$Second]) with Ticket number $($Second)"
Write-Host "Third place is $($Tickets[$Third]) with Ticket number $($Third)"
