# ==============================================================================================
#    NAME: WIC.ps1
#  AUTHOR: Paul Sarbinowski and Sebastien Dellabella
#    DATE: June 2013
# Description: Prepares a customized Windows image by cloning an initial VM, 
# automatically installing updates in it and then syspreping it to generate a .wim file.
# ==============================================================================================

# Show menu
Write-Host " " 
Write-Host "                                                   " -BackgroundColor Green
Write-Host "Welcome to the Windows Image Customisation project!" -BackgroundColor Green -ForegroundColor Black
Write-Host "                                                   " -BackgroundColor Green
Write-Host " " 
Write-Host "Please select the operating system image you want to generate" -ForegroundColor Green
Write-Host " " 
Write-Host "1. Windows 7 Enterprise with standard applications AND security updates" -ForegroundColor Magenta
Write-Host "2. Windows 7 Enterprise with security updates ONLY" -ForegroundColor Magenta
Write-Host "3. Windows 7 Enterprise X64 with standard applications AND security updates" -ForegroundColor Yellow
Write-Host "4. Windows 7 Enterprise X64 with security updates ONLY"  -ForegroundColor Yellow
Write-Host "5. Windows 8 Enterprise with standard applications AND security updates" -ForegroundColor Cyan
Write-Host "6. Windows 8 Enterprise with security updates ONLY"  -ForegroundColor Cyan
Write-Host " "
$choice = Read-Host "Select 1-6"
 
Write-Host " "
 
switch ($choice) 
    { 
        1 {
           "** Windows 7 Enterprise with standard applications AND security updates **";
           $OS = "w7ent";
           $Arch = "x86";
           $SP = "sp1";
           $LM = $false;
           break;
          } 
        2 {
           "** Windows 7 Enterprise with security updates ONLY **";
           $OS = "w7ent";
           $Arch = "x86";
           $SP = "sp1";
           $LM = $true;
           break;
          } 
        3 {
           "**  Windows 7 Enterprise X64 with standard applications AND security updates **";
           $OS = "w7ent";
           $Arch = "x64";
           $SP = "sp1";
           $LM = $false;
           break;
          } 
        4 {
           "** Windows 7 Enterprise X64 with security updates ONLY **";
           $OS = "w7ent";
           $Arch = "x64";
           $SP = "sp1";
           $LM = $true;
           break;
          } 
        5 {
           "** Windows 8 Enterprise with standard applications AND security updates **";
           $OS = "w8ent";
           $Arch = "x64";
           $SP = "rtm";
           $LM = $false;
           break;
          } 
        6 {
           "** Windows 8 Enterprise with security updates ONLY **";
           $OS = "w8ent";
           $Arch = "x64";
           $SP = "rtm";
           $LM = $true;
           break;
          } 
        default {
          "** The selection could not be determined **";
          Write-Host "Invalid choice! Exiting..." -ForegroundColor Red;
          Exit
          break;
          }
    }

$SCVMM_SnapIn   = "Microsoft.SystemCenter.VirtualMachineManager"
$VMMServerName  = "VMMSERVERADDRESS"
$OSName     = $OS + $Arch + $SP
$OriginalVMName = "cln" + $OSName
$LocalAccount   = "LOCALADMINUSERNAME"
$LocalPass  = "LOCALADMINPASSWORD"
$Owner = "OWNEROFTHEVM"
$ProdTemplatePath = "WindowsImageCustomization"
$ProdTemplateHostGroups = "HOSTGROUPFORTHETEMPLATE"
$mg = ""
$wim = ""

if ($LM -eq $true) {
    $NewVMName  = $OriginalVMName + "l"
    $ProdTemplateName = "te-" + $OSName + "#lm";
    $mg = "LM"
} else {
    $NewVMName  = $OriginalVMName + "t"
    $ProdTemplateName = "te-" + $OSName + "#cm";
    $mg = "CM"
}

$VMName = $NewVMName;
$LocalMountLetter    = "\\$VMName\C$"
$cmfPath = "$LocalMountLetter\Program Files\CERN\CMF\CMFReport.txt"

function EnableAutoLogon()
{
    Write-Host "Enabling AutoLogon" -ForegroundColor Green
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v AutoLogonCount /t REG_SZ /d 100 /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v AutoAdminLogon /t REG_SZ /d 1 /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultUserName /t REG_SZ /d $LocalAccount /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultPassword /t REG_SZ /d $LocalPass /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultDomainName /t REG_SZ /d $VMName /f", 20)
}

function DisableAutoLogon()
{
    Write-Host "Disabling AutoLogon..." -ForegroundColor Green
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v AutoLogonCount /t REG_SZ /d 0 /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v AutoAdminLogon /t REG_SZ /d 0 /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultUserName /t REG_SZ /d $LocalAccount /f", 20)
    Run_Remotely("reg.exe DELETE 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultPassword /f", 20)
    Run_Remotely("reg.exe ADD 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /v DefaultDomainName /t REG_SZ /d $VMName /f", 20)
}

function Wait_Alive()
{
    # Ping Test to see if the machine is up, wait if it isn't
    while (! (Test-Connection -ComputerName $VMName -Count 2 -Quiet) ) {
        Write-Host "$VMName is still not pingable! Waiting some more..." -ForegroundColor Red 
        Start-Sleep -s 45
    }
}

function Run_Remotely($command, $timeout = 0)
{
    $cc = 1
    while ($cc -ne 0) {
        if($timeout -ne 0) 
        {
            .\PsExec.exe -accepteula \\$ipAddr -u $LocalAccount -p $LocalPass -i -n $timeout cmd.exe /c powershell -noninteractive -command $command
        }
        else
        {
            .\PsExec.exe -accepteula \\$ipAddr -u $LocalAccount -p $LocalPass -i cmd.exe /c powershell -noninteractive -command $command
        }
        $cc = $LastExitCode
    }
}

function Restart_RemoteVM($IPorDomainName) {
    .\psshutdown.exe -accepteula \\$IPorDomainName -u $LocalAccount -p $LocalPass -f -r -t 5
}

Function Rename_wim()
{
    If (Test-Path -Path "..\install86.wim")
    {
        $wim = "install" + $OS + $Arch + $SP + $mg + "-" + (Get-Date).Day + (Get-Date).month + (Get-Date).year + ".wim"
        rename-item "..\install86.wim" $wim
        Write-Host "install86.wim has been renamed to $wim" -ForeGroundColor green
        return
    }
    else
    {
        If (Test-Path -Path "..\install64.wim")
        {
        $wim = "install" + $OS + $Arch + $SP + $mg + "-" + (Get-Date).Day + (Get-Date).month + (Get-Date).year + ".wim"
        rename-item "..\install64.wim" $wim
        Write-Host "install64.wim has been renamed to $wim" -ForeGroundColor green
        return
        }
        else
        {
        Write-Host "The .wim file doesnt exist..." -ForeGroundColor red
        }
    }
}


function Format_CSC()
{
        Write-Host "Adding  registry key to format the Offline files cache..." -ForegroundColor Green
        $regcmd = "reg.exe ADD 'HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\CSC\Parameters' /v FormatDatabase /t REG_DWORD /d 1 /f"
        Run_Remotely($regcmd, 20)
        Write-Host "Changing startup type for the Offline file service..." -ForegroundColor Green
        $regcmd = "reg.exe ADD 'HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\CSC' /v Start /t REG_DWORD /d 1 /f"
        Run_Remotely($regcmd, 20)
}

function Remove_users
{
        $Group = [ADSI]("WinNT://" + $ipAddr + "/Administrators")
        $members = @() 
        $Group.Members() | 
             % { 
                 $AdsPath = $_.GetType().InvokeMember("Adspath", 'GetProperty', $null, $_, $null) 
                 # Domain members will have an ADSPath like WinNT://DomainName/UserName. 
                 # Local accounts will have a value like WinNT://DomainName/ComputerName/UserName. 
                 $a = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries) 
                 $name = $a[-1] 
                 $domain = $a[-2] 
                $class = $_.GetType().InvokeMember("Class", 'GetProperty', $null, $_, $null) 
          
                 $member = New-Object PSObject 
                 $member | Add-Member -MemberType NoteProperty -Name "Name" -Value $name 
                 $member | Add-Member -MemberType NoteProperty -Name "Domain" -Value $domain 
                 $member | Add-Member -MemberType NoteProperty -Name "Class" -Value $class 
          
                 $members += $member 
              }

        #
        # Remove all group members except NICE Desktops Administrators, Domains Admins and local Administrator
        #
        echo ""
        Write-host "Here is the list of users currently in the Administrator group of the machine:"
        $Members
        echo ""
        foreach ($m in $members) {
               $AdminName = $m.Name
               $DomainName = $m.Domain
               if (($AdminName -eq $LocalAccount) -or ($AdminName -imatch "Domain Admins"))
               { 
                    Write-Host "Ok, $DomainName\$AdminName is on white list" -ForegroundColor Green
               }
               else 
               {
                    Write-Host "$DomainName\$AdminName will be removed" -ForegroundColor Red
                    $group.Remove("WinNT://$DomainName/$AdminName")
               }
         }
}

# Clone VM 
Write-Host "1. Creating clone from reference VM..." -ForegroundColor Green
Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue
$vmmserver = Get-VMMServer $VMMServerName

# Load old VM and stop it
$OriginalVM = Get-VM -VMMServer $vmmserver -Name $OriginalVMName
Stop-VM -VM $OriginalVM

$VMHost = $OriginalVM.VMHost
# Create a new VM based on the reference (if it already exists stop it and delete it)
$temp = Get-VM -VMMServer $vmmserver -Name $NewVMName
if ($temp) {
    Stop-VM VM $temp
    Remove-VM $temp -Force
}
New-VM -VM $OriginalVM -Name $NewVMName -VMHost $VMHost -Path $VMHost.VMPaths[0]
$NewVM = Get-VM -VMMServer $VMMServer -Name $NewVMName
Write-Host "VM Cloning done" -ForegroundColor Green

Write-Host "2. Modifying new VM to match entry in LanDB..." -ForegroundColor Green
# Modify new vm to have different mac address to avoid conflicts in network
$macAddr = "00:11:22:33:44:55" #CHANGE THIS

Write-Host "3. Create new network adapter with new MAC..." -ForegroundColor Green

# Create new network adaptor with the mac registered in landb
Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue
$vmmserver = Get-VMMServer $VMMServerName
$NewVM = Get-VM -VMMServer $VMMServer -Name $NewVMName
$NewVM = Stop-VM -VM $NewVM
Start-Sleep -s 120
$oldadapter = Get-VirtualNetworkAdapter -vm $NewVM
$PhysicalAddressType = $oldadapter.PhysicalAddressType
$PhysicalAddress = $oldadapter.PhysicalAddress
$VirtualNetwork = $oldadapter.VirtualNetwork
Remove-VirtualNetworkAdapter -VirtualNetworkAdapter $oldadapter
New-VirtualNetworkAdapter -vm $NewVM -PhysicalAddressType $PhysicalAddressType -PhysicalAddress $macAddr -VirtualNetwork "Network Name" -Synthetic 

Write-Host "Network adapter created." -ForegroundColor Green

# Rename computer
Write-Host "4. Renaming VM..." -ForegroundColor Green

$NewVM = Start-VM -VM $NewVM
Wait_Alive
Start-Sleep -s 240
$mydev = $lanDBsvc.getDeviceInfo($NewVMName)
$ipAddr = $mydev.Interfaces[0].IPAddress

$command    = "(Get-WmiObject -Class Win32_ComputerSystem).Rename('$VMName')"
Run_Remotely($command, 0)

Restart_RemoteVM($ipAddr)
Write-Host "VM Renamed." -ForegroundColor Green
Start-Sleep -s 120
Wait_Alive

Write-Host "5. Enabling AutoLogon..." -ForegroundColor Green
EnableAutoLogon
Write-Host "AutoLogon enabled..." -ForegroundColor Green
Restart_RemoteVM($ipAddr)
Start-Sleep -s 120
Wait_Alive

# Update remote machine with all the critical updates except language packs of course
$cc = 1
while ($cc -ne 0) {
    Write-Host "6. Trying to apply critical updates on VM..." -ForegroundColor Green
    .\PsExec.exe -accepteula \\$VMName -u $LocalAccount -p $LocalPass -i -n 30 cmd.exe /c 'echo . | powershell -noninteractive -command "set-executionpolicy remotesigned -force; $needsReboot = $false; $UpdateSession = New-Object -ComObject Microsoft.Update.Session; $UpdateSearcher = $UpdateSession.CreateUpdateSearcher(); Write-Host \" - Searching for Updates\"; $SearchResult = $UpdateSearcher.Search(\"IsAssigned=1 and IsHidden=0 and IsInstalled=0\"); Write-Host \" - Found [$($SearchResult.Updates.count)] Updates to Download and install\"; foreach($Update in $SearchResult.Updates) {if(!($Update.Title.ToLower() -like \"*language*\")){ $UpdatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl; if ( $Update.EulaAccepted -eq 0 ) { $Update.AcceptEula() }; $UpdatesCollection.Add($Update) | out-null; Write-Host \" + Downloading Update $($Update.Title)\"; $UpdatesDownloader = $UpdateSession.CreateUpdateDownloader(); $UpdatesDownloader.Updates = $UpdatesCollection; $DownloadResult = $UpdatesDownloader.Download(); Write-Host \" - Installing Update\"; $UpdatesInstaller = $UpdateSession.CreateUpdateInstaller(); $UpdatesInstaller.Updates = $UpdatesCollection; $InstallResult = $UpdatesInstaller.Install(); $needsReboot = $installResult.rebootRequired }}; "'
    Start-Sleep -s 60
    $cc = $LastExitCode
}
Write-Host "First set of updates applied! Restarting..." -ForegroundColor Green
Restart_RemoteVM($VMName)

Write-Host "Waiting for restart to finish..."
Start-Sleep -s 300
Wait_Alive

# Update remote machine with all of the updates except language packs
$cc = 1
while ($cc -ne 0) {
    Write-Host "7. Trying to apply all updates on VM..." -ForegroundColor Green
    .\PsExec.exe -accepteula \\$VMName -u $LocalAccount -p $LocalPass -i -n 30 cmd.exe /c 'echo . | powershell -noninteractive -command "set-executionpolicy remotesigned -force; $needsReboot = $false; $UpdateSession = New-Object -ComObject Microsoft.Update.Session; $UpdateSearcher = $UpdateSession.CreateUpdateSearcher(); Write-Host \" - Searching for Updates\"; $SearchResult = $UpdateSearcher.Search(\"IsHidden=0 and IsInstalled=0\"); Write-Host \" - Found [$($SearchResult.Updates.count)] Updates to Download and install\"; foreach($Update in $SearchResult.Updates) {if(!($Update.Title.ToLower() -like \"*language*\")){ $UpdatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl; if ( $Update.EulaAccepted -eq 0 ) { $Update.AcceptEula() }; $UpdatesCollection.Add($Update) | out-null; Write-Host \" + Downloading Update $($Update.Title)\"; $UpdatesDownloader = $UpdateSession.CreateUpdateDownloader(); $UpdatesDownloader.Updates = $UpdatesCollection; $DownloadResult = $UpdatesDownloader.Download(); Write-Host \" - Installing Update\"; $UpdatesInstaller = $UpdateSession.CreateUpdateInstaller(); $UpdatesInstaller.Updates = $UpdatesCollection; $InstallResult = $UpdatesInstaller.Install(); $needsReboot = $installResult.rebootRequired }}; "'
    Start-Sleep -s 60
    $cc = $LastExitCode
}
Write-Host "Second set of updates applied! Restarting..." -ForegroundColor Green
Restart_RemoteVM($VMName)

Write-Host "Waiting for restart action to finish..." 
Start-Sleep -s 300
Wait_Alive

# Join domain
# Ping Test to see if the machine is up, wait if it is.
while (Test-Connection -ComputerName $VMName -Count 2 -Quiet) {
    Write-Host "8. $VMName is trying to join the domain..." -ForegroundColor Green 
    # try to establish the connection with psexec several times until the result is different from 1460 - Psexec Timeout error code.
    $cc = 0
    do {
        .\PsExec.exe -accepteula \\$VMName -u $LocalAccount -p $LocalPass -s -n 20 -d cmd.exe /c 'echo . | powershell -noninteractive -command "if ($(Add-Computer -DomainName DOMAINURL -Credential $(New-Object System.Management.Automation.PSCredential 'DOMAINNAME\DOMAINACCOUNT', $(ConvertTo-SecureString  -AsPlainText 'DOMAINPASS' -Force)) -PassThru)) { Stop-Computer -Force }"' 
        $cc = $LastExitCode
    }
    while ($cc -eq 1460) 
    Start-Sleep -s 60
}
Write-Host "Domain joined!" -ForegroundColor Green
Write-Host "Restarting $NewVM" -ForegroundColor Green
$NewVM = Start-VM -VM $NewVM
Wait_Alive
Start-Sleep -s 120

# Mount C: on remote machine
Write-Host "9. Mounting $LocalMountLetter from VM..." -ForegroundColor Green
While(!(Test-Path -path $cmfPath )) {
    net use $LocalMountLetter /user:$VMName\$LocalAccount $LocalPass /persistent:yes
    Write-Host "$cmfpath is not reachable, restarting the machine..." 
    Restart_RemoteVM($VMName)
    Write-Host "Waiting for restart action to finish..." 
    Start-Sleep -s 120
}
Write-Host "$LocalMountLetter has been mounted" -ForegroundColor Green

Start-Sleep -s 5

# While CMF is still working (installing additional software) wait...
Write-Host "10. Waiting for CMF to install all apps and patches..." -ForegroundColor Green
While(!(Test-Path -path $cmfPath ) -or !(Select-String -Path $cmfPath -Pattern "EXECALLCOMPLETE=True") ) {
    Write-Host "CMF on $VMName is still working! Waiting some more..." -ForegroundColor Red 
    Start-Sleep -s 300
}
Write-Host "CMF has finished its job! Stopping cmf agent and cleaning up logs..." -ForegroundColor Green
while (((Get-WmiObject -computerName $VMName Win32_Service -Filter "Name='cmfagent'").InterrogateService().ReturnValue) -ne 6)  { 
    (Get-WmiObject -computerName $VMName Win32_Service -Filter "Name='cmfagent'").StopService().ReturnValue 
    Start-Sleep -s 5
}

# Delete CMFAgent.exe, all logs
Remove-Item "$LocalMountLetter\Program Files\CERN\CMF\Logs\*" -recurse -Force -ErrorAction SilentlyContinue

# Set EXECALLCOMPLETE=False, EXECALLPENDING=True and remove DETECTTOOLLASTRUN line from CMFReport.txt

(Get-Content $cmfPath) | Foreach-Object {$_ -replace "EXECALLCOMPLETE=True", "EXECALLCOMPLETE=False" -replace "EXECALLPENDING=False", "EXECALLPENDING=True" -replace "^DETECTTOOLLASTRUN=.+$", ""} | where {$_ -ne ""} | Set-Content $cmfPath

Write-Host "11. Disabling AutoLogon..." -ForegroundColor Green
DisableAutoLogon
Write-Host "AutoLogon Disabled..." -ForegroundColor Green

Write-Host "12. Scheduling Offline files database format..." -ForegroundColor Green
Format_CSC
Write-Host "Format scheduled for the next reboot..." -ForegroundColor Green

Write-Host "13. Removing useless users accounts" -ForegroundColor Green
remove_users
remove_users
Write-Host "Accounts deleted..." -ForegroundColor Green

# When finished start capturing
Write-Host "14. Starting the capturing process..." -ForegroundColor Green
Invoke-Expression(".\CaptureWimImage.ps1 -VMName $VMName -LocalAdmin $LocalAccount -LocalPass $LocalPass") 
Write-Host "Capturing process is over!" -ForegroundColor Green

# When the .wim file has been generated we rename it.
Write-Host "15. Renaming .wim file..." -ForegroundColor Green
Rename_wim
Write-Host "Renaming process is over!" -ForegroundColor Green

Write-Host "16. Creating Template..." -ForegroundColor Green

######## Template creation based on the created/updated VM ########
# Cleanup template before creating new one

if(($vmsvc.TemplateExists($ProdTemplateName)) -eq $true){
    $vmsvc.TemplateDeleteRequest($ProdTemplateName)
    Write-Host "Wait for old template deletion..." -ForegroundColor Green
    while (($vmsvc.TemplateExists($ProdTemplateName)) -eq $true) { Start-Sleep -s 30 }
}

Write-Host "Creating new Template $ProdTemplateName ..." -ForegroundColor Green
$vmsvc.TemplateCreateRequest($NewVMName,$ProdTemplateName,$ProdTemplateName,$ProdTemplatePath,$Owner,$ProdTemplateHostGroups)

######## End of template creation ########
Write-Host "End of template creation." -ForegroundColor Green

# Delete leftover VM
# Write-Host "Deleting leftover VM..." -ForegroundColor Green
# Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue
# $vmmserver = Get-VMMServer $VMMServerName
# $NewVM = Get-VM -VMMServer $VMMServer -Name $NewVMName
# $NewVM = Stop-VM -VM $NewVM
# $NewVM = Remove-VM $NewVM -Force
