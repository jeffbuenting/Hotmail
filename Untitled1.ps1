# Adds the base cmdlets
Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue
# Add the following if you want to do things with Update Manager
#Add-PSSnapin VMware.VumAutomation
# This script adds some helper functions and sets the appearance. You can pick and choose parts of this file for a fully custom appearance.
. 'c:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'

connect-viserver 192.168.1.10$

$esxcli = get-vmhost 192.168.1.10 | get-esxcli