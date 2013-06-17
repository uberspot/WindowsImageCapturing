#Automatic Windows Image Customization Project

The goal of the project is to automatically create customized .wim images of windows that will be deployed on other workstations.
Those images will contain all the OS security updates as well as extra selected software. 
Previously a standard installation took ~120-130 minutes while an installation from a preconfigured image takes about 25-30 minutes.

##Process flow

![Flow](https://raw.github.com/uberspot/WindowsImageCapturing/master/flow.png)

###Steps:
 
 - We create a new Virtual Machine with Windows 7 or 8 installed. This installation is not connected to the domain and it is based on the unchanged Microsoft ISO image. We avoid using an already sysprepped image because it's more likely to cause problems in the next stages of configuration.
 - Once the installation is ready we clone the VM, and keep the original one powered off for future use.
 - The VM will join the domain, update with the latest security patches and receive all the applications we distribute for a standard centrally managed machine
 - When the machine is up to date and complete, we prepare it for the generalization process (Sysprep).
 - We then reboot to the Windows Preinstallation Environment to capture the current VM installation without altering it!
 - Once it's done, we copy the resulting install.wim file to DFS and rename it properly. 

##Requirements

 If you start the project from the beginning, here is all what you need

To create the .wim image automatically you need:

 - Two Virtual Machines (32bit and 64bit) with windows 7 installed (the clean reference VMs) 
 - One custom winpe.iso disk
 - The automation scripts 

On the technicians machine you need:
 
 - PowerShell (has to have Hyper-V commands like Start-VM etc)
 - Windows Automated Installation kit (MS WAIK) IF you want to modify your winpe.iso
 - psexec.exe and psshutdown in the same folder as the PowerShell scripts used for capturing the images.
 - SCVMM Admin Console for easy VM management.

##Create the clean reference Virtual Machines
 
We need to create two virtual machines to use as a reference for the 32bit and 64bit .wim image respectively. Those machines will be clean Windows installations without any applications or patches preinstalled.
Each application and patch will be installed during the pre-capturing period by the scripts and CMF.

###Step 1
 
To create the reference VMs you have to install a clean, not edited, windows 7/8 version from an .iso.
Be careful when using an already syspreped image. You might run into "failed to parse the autounattended.xml" errors or similar during the installation of your custom image.
Create a virtual machine. Go to the properties of the machine, select CD as first in Startup option and mount iso image of the system you want to customize e.g. SW_DVD5_SA_Win_Ent_7_32BIT_English_Full_MLF_X15-70745.ISO.
(You may need to copy your ISO image to the vmm library if you use hyperv).
Make sure that you choose a minimum of 60GB space in the hard drive and at leasts 2GB of ram for each VM.
 
###Step 2
 
Follow the installation steps and finish the installation in your virtual machine.
 
- Create a user named whatever you like (it will be deleted later)
- Give a name to the computer that is the same as the virtual machine name (e.g. clnw7entx64sp1)
- When asked, leave the language settings to the default (English-International). Also when asked about updates click Ask me later..

###Step3

On each VM now enable the default Administrator user ( http://technet.microsoft.com/en-us/library/dd744293%28v=ws.10%29.aspx and http://www.ghacks.net/2012/06/11/windows-8-enable-the-hidden-administrator-account/ for windows 8 ), add a password for it and then login as Administrator and delete all other local accounts and their files. Only the Administrator account should be present on the system.
 
###Step 4
 
Enable autologon for Administrator so that cmf will be able to reboot and continue its work without intervention.  Just type Run in the Start button, then 'control userpasswords2' un-tick the 'users require password to login', enter the 'Administrator' accounts password  and ok.
 
###Step 5

Deactivate the Firewall via Start/Control Panel so that it doesn't block connections from psExec.
 
###Step 6

Add the following registry key to enable administrative shares via regedit.exe
 
    Hive: HKEY_LOCAL_MACHINE
    Key: Software\Microsoft\Windows\CurrentVersion\Policies\System
    Name: LocalAccountTokenFilterPolicy
    Data Type: REG_DWORD
    Value: 1
 
[http://en.wikipedia.org/wiki/Administrative_share](Administrative shares)
 
 
###Step 7

Go to Control Panel/Windows update/Let me choose my settings and choose 'never check for updates' in the dropdown selection.
After that, return to the main windows update windows and click the 'check for updates..' button. It will display a message saying 'To check for updates you must first install... ' Click 'Install now' and then Restart the machine.
This small update is necessary so that the machine can be updated later on automatically.
After that stop the VM either via the Virtual Machine Manager or via powershell.
 
###Step 8

Add the winpe.iso you will create in the next section as the source in the dvd drive of the VM via the Virtual Machine Manager (from the VM properties).
 
###Step 9

Change the VMs boot order to ("PxeBoot", "IdeHardDrive", "CD", "Floppy") with the following commands in powershell.
 
    # VM setup
    $SCVMM_SnapIn        = "Microsoft.SystemCenter.VirtualMachineManager"
    $VMMServerName       = "HYPEVSERVER"
    $VMName     = "VMNAME" # CHANGE VM NAME
    # Register the snapin if it is not registered yet
    Add-PSSnapIn $SCVMM_SnapIn -ErrorAction SilentlyContinue
    # Load the vm
    Write-Host "Loading and starting vm..." -ForegroundColor Green
    $vmmserver = Get-VMMServer $VMMServerName
    $vm = Get-VM -VMMServer $vmmserver -Name $VMName
    $vm = Stop-VM .VM $vm
    Set-VM $vm -BootOrder @("PxeBoot","IdeHardDrive","CD","Floppy")
    Write-Host "Boot order changed.." -ForegroundColor Green
 
 
*Note: Take care NOT to use more than 15 characters for the Machine name because  that is the limit for Windows Hostnames.*
 
###Step 11

Create a snapshot from your reference VM just in case via Virtual Machine Manager.
And the VM is now ready!

##Custom WinPE image
 
###Step 1

To create the custom winpe.iso disk you have to install MS WAIK and open the Start/All Programs/MS Windows AIK/ Deployment tools command prompt  in your own machine (not in the VMs) and execute:

    copype.cmd x86 c:\winpe
 
The script creates the following directory structure and copies all the necessary files for that architecture. For example,
 
    \winpe
    \winpe\ISO
    \winpe\mount

*Note: Delete the file  \ISO\boot\bootfix.bin so that it does not display a "Press a key to boot from cd..." prompt on each boot.*

###Step 2

Now you have to add:
 
 - a 32bit shutdown.exe executable taken from a standard win7 32bit installation
 - Imagex.exe executable downloaded online
 - The script that will run after the boot.
 
From the Deployment tools cmd opened before, in your machine, again do:
 
    imagex /info C:\winpe\ISO\sources\boot.wim                # check the image index (e.g. 1)
    imagex /mountrw C:\winpe\ISO\sources\boot.wim XX C:\winpe\ISO\sources\boot\  #instead of XX use the image index seen in the previous command
 
###Step 3

Now you have a mounted boot.wim in C:\winpe\ISO\sources\boot:

 - Copy the shutdown.exe and imagex.exe to C:\winpe\ISO\sources\boot\Windows\System32\ 
 - Edit C:\winpe\ISO\sources\boot\Windows\System32\startnet.cmd.
 
Add the following commands in startnet.cmd:
 
    # 1) Winpe starts and mounts dfs directory to save image there
    net use z: \\YOURDFSSERVERHERE\dfs /user:DOMAINNAME\DOMAINUSER DOMAINPASSWORD
 
    # 2) Capture image to somedir
    if exist "C:\Program Files(x86)" (imagex.exe /compress fast /capture c:\ z:\WindowsImageCustomization\install64.wim "new64imagefile") else ( imagex.exe /compress fast /capture d:\ z:\WindowsImageCustomization\install86.wim "new86imagefile" )
 
    # 3) Shutdown pc after capturing
    shutdown.exe .s
 
Note: The "if exist C:\Program Files(x86)" command checks to see if that folder exists. If it exists then we presume that we are in the 64bit virtual machine and we name the resulting .wim image accordingly.
 
###Step 4

Finally to write the changes we made to our winpe image we must do from the Deployment tools cmd, in your machine:
 
    imagex /unmount C:\winpe\ISO\sources\boot\ /commit
    oscdimg -n -bc:\winpe\etfsboot.com c:\winpe\ISO c:\winpe\winpe.iso
 
To create a bootable C:\winpe\winpe.iso that's ready for usage as an extra CD Drive in each VM.
Now you can use it in the VMs. 

##Capturing process/Usage of scripts
 
###Step 1

 Edit the WindowsImageCapturing.ps1 script that starts the cloning/updating/capturing process and change the nesessary variables in the beginning.
 
*Note: After you run it the whole procedure should last about 5-6 hours.*
 
The script will capture the image to \\YOURDFSSERVER\WindowsImageCustomization\install86(or64).wim and it should be about 6-8GB in size depending on the updates installed.

*Note:
 
 - If anything goes wrong during the process and the image creation fails you can easily restart by running the script again. 
   It will automatically delete the previous copied VM and start over.
- The above scripts require PsExec.exe, PsShutdown.exe AND CaptureWimImage.ps1 to be in the same directory in your machine as the PowerShell script to run properly.
- If during the capturing the scripts get stuck during the PsExec phase and psexec seems too not respond you have to go to each VM and do the following steps:

To fix this issue you will need to:

 - Delete the c:\windows\psexesvc.exe file from the vm machine.
 - Launch Regedit
 - Navigate to the key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\PSEXESVC
 - Modify the value of "ImagePath" to read "%SystemRoot%\PSEXESVC.EXE"

The next time you use psexec with this machine it should run fine.

All that needs to be done after the scripts end is to copy the .wim file to the \sources\install.wim of our windows installation folder.

##Create a windows .iso for installations
 
To create an .iso for installations of any kind you simply run as an Administrator the Start/All Programs/Microsoft Windows AIK/Deployment Tools Command Prompt and you execute the following command.

    oscdimg -bC:\WindowsInstallationFolder\boot\etfsboot.com -u2 -h -m -lWin_dvd C:\WindowsInstallationFolder C:\WindowsImageCustomization\finalimage.iso

assuming that you want to make an iso from the installation files located in e.g. C:\WindowsInstallationFolder

Note: If etfsboot.com is not located in \boot and it's located in the root directory of the installation files simply adjust the first directory in the parameters. 

##License

This work is licensed under a [http://creativecommons.org/licenses/by/3.0/deed.en_US](Creative Commons Attribution 3.0 Unported License)
