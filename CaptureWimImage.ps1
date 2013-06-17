# This script is used to load, start, sysprep, capture the image and then restore the VM
# Example usage: .\CaptureWimImage.ps1 -VMName $VMName -LocalAdmin $LocalAccount -LocalPass $LocalPass

param(
      # The hostname of the VM which is to be captured
      [parameter(Mandatory=$true)] [string] $VMName,
      # The local account username with administrative privileges or domain admin username
      [parameter(Mandatory=$true)] [string] $LocalAdmin,
      # The local admin password
      [parameter(Mandatory=$true)] [string] $LocalPass
)

# 1) VM setup 
$SCVMM_SnapIn        = "Microsoft.SystemCenter.VirtualMachineManager"
$VMMServerName       = "URL OF HYPERV VM SERVER"


# Ping Test to see if the machine is up, exit if it isn't
if (Test-Connection -ComputerName $VMName -Count 4 -Quiet) {
        Write-Host "$VMName is alive and pinging!" -ForegroundColor Green 
} else { 
        Write-Host "$VMName seems dead. Not responding to ping. Exiting..." -ForegroundColor Red
        Exit
}

# Register the snapin if it is not registered yet
Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue

# Load the vm (and start it as well)
Write-Host "Loading and starting VM..." -ForegroundColor Green 

$vmmserver = Get-VMMServer $VMMServerName
$vm = Get-VM -VMMServer $vmmserver -Name $VMName

$vm = Start-VM –VM $vm

# 2) Create vm snapshot and clean up older snapshots to save space
Write-Host "Creating new vm snapshot..." -ForegroundColor Green 
$LastCheckpoint = Get-VMCheckpoint -MostRecent -vm $vm 
if(!($LastCheckpoint -eq $null)) { Remove-VMCheckpoint -VMCheckpoint $LastCheckpoint }
New-VMCheckpoint $vm

Start-Sleep -s 30

# 3) Connect to the vm with pstools, start sysprep (machine shuts down) 

Write-Host "Trying to SYSPREP the VM..." -ForegroundColor Green

# Ping Test to see if the machine is up, wait if it is.
while (Test-Connection -ComputerName $VMName -Count 2 -Quiet) {
        Write-Host "$VMName is still pingable! Trying to SYSPREP..." -ForegroundColor Red 
        # We try to establish the connection with psexec several times until the result is different from 1460 - Psexec Timeout error code.
        $cc = 0
        do {
            .\PsExec.exe -accepteula \\$VMName -u $LocalAdmin -p $LocalPass -s -n 20 -d cmd.exe /c "C:\Windows\System32\sysprep\sysprep.exe /generalize /oobe /shutdown"    
            $cc = $LastExitCode
        }
        while ($cc -eq 1460) 

        Start-Sleep -s 600
}

Write-Host "SYSPREP has been successfully applied!" -ForegroundColor Green

# 4) Change the boot order of the vm to boot to CD
Write-Host "Rebooting to cd..." -ForegroundColor Green 

$vm = Stop-VM –VM $vm
Set-VM $vm -BootOrder @("PxeBoot","CD","IdeHardDrive","Floppy")
$vm = Start-VM –VM $vm

Write-Host "Step 1 completed! Waiting for ~1 hour till the capturing of the image completes and then running step 2..." -ForegroundColor Green 
Start-Sleep -s 3600

# Check if the capturing has finished and the VM is powered off
while($vm.status -ne "PowerOff") {
    Write-Host "Still capturing. Waiting till machine powers off..." -ForegroundColor Red 
    Start-Sleep -s 300
}

# Register the snapin if it is not registered yet
Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue

# Load the vm (and start it as well)
Write-Host "Loading vm..." -ForegroundColor Green 
$vmmserver = Get-VMMServer $VMMServerName
$vm = Get-VM -VMMServer $vmmserver -Name $VMName

# 2) Restore previous (pre-sysprep) snapshot
Write-Host "Restoring previous snapshot..." -ForegroundColor Green 
$Checkpoint = Get-VMCheckpoint -MostRecent -vm $vm 
Restore-VMCheckpoint -VMCheckpoint $Checkpoint

# 3) Change the boot order of the vm to boot to Hard Disk
Write-Host "Changing boot order to hard drive..." -ForegroundColor Green 
$vm = Stop-VM –VM $vm
Set-VM $vm -BootOrder @("PxeBoot","IdeHardDrive","CD","Floppy")
$vm = Start-VM –VM $vm

Write-Host "Step 2 completed! Capturing complete!" -ForegroundColor Green 

