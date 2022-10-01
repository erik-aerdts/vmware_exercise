############ Ownership and Copyright
# Script created by Erik Aerdts May 2021
# modified on 29 sept 2021 by Jan van Rooij
# Copyright notice: This script can and may be modified, at your own convenience, 
# as long as the original header is preserved.
# 
# Warranty and Liability dismissal notice: 
# Erik and Jan will and can not be held responsible nor accountable for any damage,
# following improper use, nor will they be held liable for any damage directly,
# or indirectly caused as a result from using this script, either original or modified.
#
# This script demos: 
# 1) automated connection to Vcenter 
# 2) edit your VM Notes by prompt
# 3) delete snapshots older than...
# 4) list VM parameters and save in a XLS
#
############ Setup presconditions
# First, create vcenter.cred (needed once) by: 
# $vccred = Get-Credential
# 
# Second, save the credentials in a XML:
# $vccred | Export-Clixml -Path C:\scripts\vcenter.cred
# 
# Next: Suppress Certificate invalid error (needed once):
# Set-PowerCLIConfiguration -InvalidCertificateAction Ignore


############# Connect to VIcenter
$vccred = Import-Clixml -Path /Users/erikaerdts/Documents/scripts/vcenter.cred
Connect-VIServer vcenter.fhict.local -Credential $vccred


############# Get all VM's from Jan van Rooij, based on the equivalent name of the folder (!)
$servers = get-vm -Location "Erik Aerdts"

############# Create notes in VM's by inputprompt
foreach ($server in $servers) {
                            write-host "$server.name notitie: "-NoNewline ; $notitie = Read-Host
                            #set-vm $server -Notes $notitie -Confirm:$false
                            }


############# Delete snaps more recent than 10 minutes
foreach ($server in $servers) {
  $oldsnaps = get-vm $server | Get-Snapshot |where {$_.Created -gt (Get-Date).Addminutes(-10)} # | Remove-Snapshot #| select -ExpandProperty created 
                          
            foreach ($snap in $oldsnaps) { Write-Host "snapshot $snap.name on $server too young, deleting....."
                                          #Remove-Snapshot -snapshot $snap  -Confirm:$false  -WhatIf}
                                          }

############# Log all Erik's VM's specific settings in an Excel spreadsheet
# Open Excel
$excel=New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$ws=$workbook.WorkSheets.item(1)
$ws1=$workbook.WorkSheets(1)
$excel.Visible = $True
$MissingType = [System.Type]::Missing
$regel=1
# create headline
$ws.Cells.item($regel,1).value2="vcenternaam"
$ws.Cells.item($regel,2).value2="dnsnaam"
$ws.Cells.item($regel,3).value2="status"
$ws.Cells.item($regel,4).value2="datastore"
$ws.Cells.item($regel,5).value2="disks"
$ws.Cells.item($regel,6).value2="ip"
$ws.Cells.item($regel,7).value2="cpu"
$ws.Cells.item($regel,8).value2="mem"

$regel=2
$VMsAdv = Get-VM -Location "Erik Aerdts"| Sort-Object Name | % { Get-View $_.ID } 
$results=@()
ForEach ($VMAdv in $VMsAdv) 
{ 
# datastore 
$dsid=$vmadv.datastore.type+'-'+$vmadv.datastore.value
$datastore=Get-Datastore -Id $dsid | Select -ExpandProperty Name
$servernaam = $vmadv.Guest.HostName
$vmips=(Get-VM -Name $VMAdv.Name).Guest.IPAddress
# os
$os = Get-VMGuest $vmadv.Name | select osfullname
# IP
$ipall = Get-VMGuest -VM $vmadv.name | select -ExpandProperty IPAddress | Select-String -List 1
$ip = $ipall[0]
# mem&cpu
$cpu = (Get-VM -Name $VMAdv.Name).numcpu
$mem = (Get-VM -Name $VMAdv.Name).memorymb
# VMDK
$disks=$null
$naam =$vmadv.name.Split()[0]

    ForEach ($Disk in $VMAdv.Layout.Disk) 
    {   $aantal = $vmadv.Layout.Disk.count
        if ( $aantal -gt 1)
           { $diskname = $disk.DiskFile.split()[1]
             $disks += $diskname + "," 
             $vmdk=$disks -Split(','),2 
             $vmdk = $vmdk[1]
            }
    }
         if ( $aantal -eq 1) { $vmdk="geen extra vmdk" }    
$ws.Cells.item($regel,1).value2=$vmadv.name
$ws.Cells.item($regel,2).value2=$vmadv.Guest.HostName
$ws.Cells.item($regel,3).value2=$vmadv.Guest.GuestState 
$ws.Cells.item($regel,4).value2="$datastore"
$ws.Cells.item($regel,5).value2="$vmdk"
$ws.Cells.item($regel,6).value2="$ip"
$ws.Cells.item($regel,7).value2="$cpu"
$ws.Cells.item($regel,8).value2="$mem"
   
   $regel = $regel+1 
   
   }        
 
 $workbook.SaveAs('/Users/erikaerdts/Documents/serverinfo.xlsx')
 $excel.Quit()
