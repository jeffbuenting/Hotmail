# Adds the base cmdlets
Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue
# Add the following if you want to do things with Update Manager
#Add-PSSnapin VMware.VumAutomation
# This script adds some helper functions and sets the appearance. You can pick and choose parts of this file for a fully custom appearance.
. 'c:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'

$HostName = '192.168.1.10'

Connect-VIServer -Server 192.168.1.10 -User root -Password Branman1!

#get-vmhost -Name $HostName | set-VMHost -State Maintenance

( get-vm -Location $hostName | where PowerState -eq 'PoweredOn' ).Count
( get-vm -Location $hostName ).Count

#Get-VMHost -Name $HostName