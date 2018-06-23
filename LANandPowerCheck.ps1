#########################################
#                                       #
# LAN, Power, and VPN Status v1.0.A     #
# Created by: Noah Dillon & Ben Tefend  #
# Created on: 06/20/2018                #
# Revised on: 06/21/2018                #
#                                       #######################################################
#                                                                                             #
# This script was created as a pre-check for Windows 10 Upgrades during an OSD task sequence. #
# Our script references status: Up vs. Disconnected vs. Disabled.                              #
#                                                                                             #
# In order to to leverage windows forms / message boxes during a task sequence, you need to   #
# add ServiceUI.exe and TSProgressUI.exe (both x86) into your package source. You can locate  #
# these executibles on your primary site server.                                              #
#                                                                                             #
# When you add this to a task sequence, you will create a CMD step and run the following      #
# command (be sure to include your package source):                                           ###########################################################################################
#                                                                                                                                                                                       #
# ServiceUI.exe -process:TSProgressUI.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File LANandPowerCheck.ps1 #
#                                                                                                                                                                                       #
# For reference:                                                                              ###########################################################################################
# https://modalyitblog.wordpress.com/2016/10/03/powershell-gui-reboot-prompt/                 #
#                                                                                             #
# Use below command to gather list of network adapters:                                       #
#                                                                                             #
# Get-NetAdapter | select Name, PhysicalMediaType, Status                                     #
#                                                                                             #
###############################################################################################

###############################################################################################
# Hide Console Window - We are hiding this in the task sequence with the serviceui.exe (x86). #
# You can uncomment this portion during manual testing.                                       #
###############################################################################################

#Add-Type -Name Window -Namespace Console -MemberDefinition '
#[DllImport("Kernel32.dll")]
#public static extern IntPtr GetConsoleWindow();
 
#[DllImport("user32.dll")]
#public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
#'

#$consolePtr = [Console.Window]::GetConsoleWindow()
#0=Hide 5=Show
#[Console.Window]::ShowWindow($consolePtr, 0)

####################################################################################
# Hide the task sequence progress dialog so that the message boxes are noticeable. #
####################################################################################

$TSProgressUI = new-object -comobject Microsoft.SMS.TSProgressUI
$TSProgressUI.CloseProgressDialog()

##############################################################################
# Variables - Change 802.3 to whatever PhysicalMediaType you're referencing. #
# This can be updated to use Name, etc.                                      #
##############################################################################

$netadapter = Get-NetAdapter | Where-Object PhysicalMediaType -EQ "802.3"
$adapter = Get-NetAdapter | Where-Object PhysicalMediaType -EQ "802.3"
$Power = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ComputerName $env:COMPUTERNAME).PowerOnLine

################################################################
# Add Assembly "System.Windows.Forms" to create message boxes. #
################################################################

Add-Type -AssemblyName System.Windows.Forms | Out-Null

#######################################################################################################
# Information window to tell user to not run this on a docking station. This was added because we are #
# leveraging this script in a OSD task sequence. It can be removed/modified.                          #
#######################################################################################################

$warning = [System.Windows.Forms.MessageBox]::Show("If you are docked, please undock and connect both LAN and POWER to your computer directly. It's likely you'll need to restart the upgrade process after undocking. This will take a while. Hit OK to continue!","WINDOWS 10 UPGRADE PRE-CHECK",
[System.Windows.Forms.MessageBoxButtons]::OK)
#$warning.TopMost = $True

#############################################################################
# Check if user is connected to VPN. If connected, exit script with code 1. #
#############################################################################

function VPNCheck{
$vpnadapter = Get-NetAdapter | Where-Object Name -EQ "Ethernet 2"

if($vpnadapter.status -eq "Up"){
[System.Windows.Forms.MessageBox]::Show("You are currently connected to VPN. Please disconnect, plug in both LAN and POWER to your computer directly, and restart the task sequence.", "VPN WARNING")
Exit 1
}
}

#######################################################################################################
# Begin checks. It will first check if the network adapter equals up then check if power equals true. #
# Depending on the results, it will prompt the user and repeat the checks until true. Function can be #
# canceled by hitting the cancel button. Script will exit with 0 if success, 1 if failure.            #
#                                                                                                     #
# Original CheckCable script resource:                                                                #
# http://guidestomicrosoft.com/2016/07/23/check-if-the-network-cable-is-unplugged-through-powershell/ #
#######################################################################################################

function CheckCable{

    param(
    [parameter(Mandatory=$true)]
    $adapter
    )

    if($adapter.status -ne "Up" -and $Power -ne $true){
    do{
            #User is not on LAN or Power
            $Error1 = [System.Windows.Forms.MessageBox]::Show("LAN and power is not plugged in. Please connect both LAN and POWER! Hit OK to continue!","WINDOWS 10 UPGRADE PRE-CHECK",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel)
            If ($Error1 -imatch "Cancel") {
            [System.Windows.Forms.MessageBox]::Show("Upgrade has been canceled. Please run the Upgrade Sequence again. You may need to give your computer some time to retrieve policy before you can retry successfully.","UPGRADE CANCELED",
            [System.Windows.Forms.MessageBoxButtons]::OK)
            Exit 1
            } 
            $Power = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ComputerName $env:COMPUTERNAME).PowerOnLine
            Start-Sleep -Seconds 3
        }
        until($Power -eq $true -and $adapter.status -eq "Up" -or $Error1 -imatch "Cancel")
        }
    ElseIf($adapter.status -eq "Up" -and $Power -eq $true){
    do{
            #User is on LAN and Power
            $Error1 = [System.Windows.Forms.MessageBox]::Show("LAN and POWER are plugged in. Hit OK to continue!","WINDOWS 10 UPGRADE PRE-CHECK",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel)
            If ($Error1 -imatch "Cancel") {
            [System.Windows.Forms.MessageBox]::Show("Upgrade has been canceled. Please run the Upgrade Sequence again. You may need to give your computer some time to retrieve policy before you can retry successfully.","UPGRADE CANCELED",
            [System.Windows.Forms.MessageBoxButtons]::OK)
            Exit 1
            } 
            $Power = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ComputerName $env:COMPUTERNAME).PowerOnLine
            Start-Sleep -Seconds 3
        }
        until($Power -eq $true -and $adapter.status -eq "Up" -or $Error1 -imatch "Cancel")
        }
    ElseIf($adapter.status -ne "Up" -and $Power -eq $true){
    do{
            #User is not on LAN but has Power plugged in
            $Error1 = [System.Windows.Forms.MessageBox]::Show("POWER is connected but LAN is not. Please connect LAN! Hit OK to continue!","WINDOWS 10 UPGRADE PRE-CHECK",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel)
            If ($Error1 -imatch "Cancel") {
            [System.Windows.Forms.MessageBox]::Show("Upgrade has been canceled. Please run the Upgrade Sequence again. You may need to give your computer some time to retrieve policy before you can retry successfully.","UPGRADE CANCELED",
            [System.Windows.Forms.MessageBoxButtons]::OK)
            Exit 1
            } 
            $Power = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ComputerName $env:COMPUTERNAME).PowerOnLine
            Start-Sleep -Seconds 3
        }
        until($Power -eq $true -and $adapter.status -eq "Up" -or $Error1 -imatch "Cancel")
        }
    ElseIf($adapter.status -eq "Up" -and $Power -ne $true){
    do{
            #User is on LAN but Power is not
            $Error1 = [System.Windows.Forms.MessageBox]::Show("LAN is plugged in but POWER is not. Please connect POWER! Hit OK to continue!","WINDOWS 10 UPGRADE PRE-CHECK",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel)
            If ($Error1 -imatch "Cancel") {
            [System.Windows.Forms.MessageBox]::Show("Upgrade has been canceled. Please run the Upgrade Sequence again. You may need to give your computer some time to retrieve policy before you can retry successfully.","UPGRADE CANCELED",
            [System.Windows.Forms.MessageBoxButtons]::OK)
            Exit 1
            } 
            $Power = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi -ComputerName $env:COMPUTERNAME).PowerOnLine
            Start-Sleep -Seconds 3
        }
        until($Power -eq $true -and $adapter.status -eq "Up" -or $Error1 -imatch "Cancel")
        }
        #Shouldnt get here...
        Else {
            [System.Windows.Forms.MessageBox]::Show("Something went wrong. Please contact the service desk to have the situation corrected.", "MISCELLANEOUS ERROR")  
            Exit 1
        }
Exit 0 
}



VPNCheck -adapter $vpnadapter
CheckCable -adapter $netadapter
