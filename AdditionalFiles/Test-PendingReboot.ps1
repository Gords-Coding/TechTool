function Test-PendingReboot
{
    #Check CBS Registry
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) 
    { 
        return $true 
    }
    #Check Windows Update
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) 
    { 
        return $true 
    }
    #Check PendingFileRenameOperations
    if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore) 
    { 
        return $true 
    }
    #Check SCCM Client
    try { 
        $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
        $status = $util.DetermineIfRebootPending()
        if(($status -ne $null) -and $status.RebootPending)
        {
            return $true
        }
    }catch{}
    return $false
}
Invoke-Command -ComputerName $args[0] ${function:Test-PendingReboot}