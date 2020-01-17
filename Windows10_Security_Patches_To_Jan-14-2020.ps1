
# Powershell script to patch Curveball and all previously patched vulnerabilities on Windows 10 1607-1909 
# as this will identify the OS & Build then download the matching patches released up to Jan 14th 2020

# Creating workspace in C:\Support for logs and files
if (![System.IO.Directory]::Exists("C:\support\")) {
    [system.io.directory]::CreateDirectory("c:\Support")
}
if (![System.IO.Directory]::Exists("C:\support\Updates")) { 
    [system.io.directory]::CreateDirectory("c:\Support\Updates")
}


# Check Bitness of OS
if ((Get-WmiObject win32_operatingsystem | Select-Object osarchitecture).osarchitecture -eq "64-bit") {
    #64 bit logic here
    Write-Output "64-bit OS Detected" | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append
    $Script:BIT = "64"
}
else {
    #32 bit logic here
    Write-Output "32-bit OS Detected" | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append
    $Script:BIT = "86"

}


# Check if windows 10 before continuing to patch
Set-Variable -Name "WINVS" -Value ([Environment]::OSVersion.Version).Major

# Logging OS Version info
$OS | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt
Get-HotFix | Select-Object HotFixID, InstalledOn | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append


# Email settings to be passed into further events
$Script:From = "your@gmail.com"
$Script:To = "your@email.com"
$Script:Cc = "other@gmail.com"
[array]$Script:Attachment = Get-ChildItem "C:\Support\Updates\*" -Include *.txt,*.evtx
$Script:Subject1 = "$env.computername does not meet the OS requirements: $OSx$BIT"
$Script:Subject2 = "$env.computername is detected as having $KB installed on $OSx$BIT"
$Script:Subject3 = "$env.computername has had $KBL installed on $OSx$BIT"
$Script:Subject3 = "$env.computername has had $KBL identified and exited. $OSx$BIT"
$Script:Body = "Logs attached related to system patching event"
$Script:SMTPServer = "smtp.gmail.com"
$Script:SMTPPort = "587"
$Script:Secret = "PASSWORD"

# Copy the line below into the email sections so it can send based on the Subject / Body #s  Uncomment afterword
# Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess


# Exit if OS is not Win10  Can remove email section if not needed
If (!($WINVS -eq 10)) {
    #Pending further email info
    Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject1 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess
    Exit
}


#Matching OS Build to a KB Identifier & Downloads
switch ((Get-CimInstance Win32_OperatingSystem).BuildNumber) {
    ############################Windows 7 & 8 (Not support in this cumulative update)############################
    7600 { $Script:OS = "W2K8R2" }
    7601 { $Script:OS = "W2K8R2SP1" }    
    9200 { $Script:OS = "W2K12" }
    9600 { $Script:OS = "W2K12R2" }
    ############################ Win 10 versions Below ############################
    ############################# 1507 ############################
    10240 { $Script:OS = "W2K10v1507" }
    10240 { $Script:KB = "KB4534296" }
    10240 { $Script:KB64 = "kb4534306-x64_fe79ab28516198be477c18e53390f802588bca6c.msu" }
    10240 { $Script:KB86 = "kb4534306-x86_b79d87ac2c7692f2c152122c818baed709e4f11e.msu" }
    #1511 (Not Supported in this cumulative update EOL, update to a higher Windows version)
    10586 { $Script:OS = "Not Supported" }
    ############################ 1607 ############################
    14393 { $Script:OS = "W2K10v1607" }
    14393 { $Script:KB = "KB4534271" }
    14393 { $Script:KB64 = "kb4534271-x64_a009e866038836e277b167c85c58bbf1e0cc5dc8.msu" }
    14393 { $Script:KB86 = "kb4534271-x86_1401cdaf3781a6b032b558afd90fff6faa5569d3.msu" }
    ############################ 1703 ############################
    15063 { $Script:OS = "W2K10v1703" }
    15063 { $Script:KB = "KB4534296" }
    15063 { $Script:KB64 = "kb4534296-x64_be00b82daf47109410f688cd3776f2b1e3e091b1.msu" }
    15063 { $Script:KB86 = "kb4534296-x86_f5cabb7aecb6000d39d0d82f84579a03ffa79f52.msu" }
    ############################ 1709 ############################
    16299 { $Script:OS = "W2K10v1709" }
    16299 { $Script:KB = "KB4534276" }
    16299 { $Script:KB64 = "kb4534276-x64_0ef8f7279a888b2243bf02a1a4a8ab92fac5131f.msu" }
    16299 { $Script:KB86 = "kb4534276-x86_accaa8ee3113a2b42a3da387c13ab59a59688bd3.msu" }
    ############################ 1803 ############################
    17134 { $Script:OS = "W2K10v1803" }
    17134 { $Script:KB = "KB4534293" }
    17134 { $Script:KB64 = "kb4534293-x64_af7ad26642b7c49788d70902f1918b9b234292cf.msu" }
    17134 { $Script:KB86 = "kb4534293-x86_eea3d9658ebced3365ba55a6cc3de62a2a67ef93.msu" }
    ############################ 1809 ############################
    17763 { $Script:OS = "W2K10v1809" }
    17763 { $Script:KB = "KB4534273" }
    17763 { $Script:KB64 = "kb4534273-x64_74bf76bc5a941bbbd0052caf5c3f956867e1de38.msu" }
    17763 { $Script:KB86 = "kb4534273-x86_26cc081cc59d5df33392e478cf9a60f8b418fd05.msu" }
    ############################ 1903 ############################
    18362 { $Script:OS = "W2K10v1903" }
    18362 { $Script:KB = "KB4528760" }
    18362 { $Script:KB64 = "kb4528760-x64_4ea59b94564145460ab025616ff10460bb7894d8.msu" }
    18362 { $Script:KB86 = "kb4528760-x86_e8a6aae0076403e9d8d68c3ccc3f753728298b83.msu" }
    ############################ 1909 ############################
    18363 { $Script:OS = "W2K10v1909" }
    18363 { $Script:KB = "KB4528760" } 
    18363 { $Script:KB64 = "kb4528760-x64_4ea59b94564145460ab025616ff10460bb7894d8.msu" }
    18363 { $Script:KB86 = "kb4528760-x86_e8a6aae0076403e9d8d68c3ccc3f753728298b83.msu" }
    ############################ 2004 (Not Released) ############################
    19041 { $Script:OS = "Not Supported" }
    19041 { $Script:KB64 = "W2K10v2004" }
    19041 { $Script:KB86 = "W2K10v2004" }
    # PENDING: Server 2016 (Uses same as 1903 & 1909) Dont use on a server OS
    12345 { $Script:OS = "Not Supported" }
    12345 { $Script:KB = "KB4528760" } 
    12345 { $Script:KB64 = "kb4528760-x64_4ea59b94564145460ab025616ff10460bb7894d8.msu" }
    12345 { $Script:KB86 = "kb4528760-x86_e8a6aae0076403e9d8d68c3ccc3f753728298b83.msu" }
    default { $Script:OS = "Not Supported" }
}


# Choose correct Bitness KB link, If Matching KB installed, end script and update by email
switch ($BIT) {
    64 { $KBL = "$KB64" }
    86 { $KBL = "$KB86" }
    default { $KBL = "Other" }
}

$Patch = Get-Hotfix -id "$KB"

if ($null -ne $Patch) {
    #System replied with info meaning the patch is installed
    Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject2 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess

}
else {
    #Shorten website
    $Script:Web = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/01/windows10.0-"
    #Create Variable that is Script wide for the download link
    $Script:LINK = "$Web$KBL"
}


If ($KBL -eq "Other") {
    
    #If the KB doesnt match or OS not supported bail out here and send alert email
    Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject4 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess

}
     
        
            

#Alert user that system is currently updating
msg console /server:localhost "Hello , your computer needs security patches and will need to remain powered on until complete. You can continue using the system but prepare by saving data as there will be a need to reboot later"



#Download the matching Cumulative Patches for the detected Windows 10 OS
if (Test-Path "C:\Support\Updates" -Pathtype Container) {
    Set-Variable -Name "URL" -Value "$LINK"
    Set-Variable -Name "Filename1" -Value "C:\Support\Updates\$KBx$BIT.tmp"
            
    function Get-FileFromURL {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, Position = 0)]
            [System.Uri]$URL,
            [Parameter(Mandatory, Position = 1)]
            [string]$Filename1
        )
    
        process {
            try {
                $request = [System.Net.HttpWebRequest]::Create($URL)
                $request.set_Timeout(5000) # 5 second timeout
                $response = $request.GetResponse()
                $total_bytes = $response.ContentLength
                $response_stream = $response.GetResponseStream()
    
                try {
                    # 256KB works better on my machine for 1GB and 10GB files
                    # See https://www.microsoft.com/en-us/research/wp-content/uploads/2004/12/tr-2004-136.pdf
                    # Cf. https://stackoverflow.com/a/3034155/10504393
                    $buffer = New-Object -TypeName byte[] -ArgumentList 256KB
                    $target_stream = [System.IO.File]::Create($Filename1)
    
                    $timer = New-Object -TypeName timers.timer
                    $timer.Interval = 1000 # Update progress every second
                    $timer_event = Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
                        $Global:update_progress = $true
                    }
                    $timer.Start()
    
                    do {
                        $count = $response_stream.Read($buffer, 0, $buffer.length)
                        $target_stream.Write($buffer, 0, $count)
                        $downloaded_bytes = $downloaded_bytes + $count
    
                        if ($Global:update_progress) {
                            $percent = $downloaded_bytes / $total_bytes
                            $status = @{
                                completed  = "{0,6:p2} Completed" -f $percent
                                downloaded = "{0:n0} MB of {1:n0} MB" -f ($downloaded_bytes / 1MB), ($total_bytes / 1MB)
                                speed      = "{0,7:n0} KB/s" -f (($downloaded_bytes - $prev_downloaded_bytes) / 1KB)
                                eta        = "eta {0:hh\:mm\:ss}" -f (New-TimeSpan -Seconds (($total_bytes - $downloaded_bytes) / ($downloaded_bytes - $prev_downloaded_bytes)))
                            }
                            $progress_args = @{
                                Activity        = "Downloading $URL"
                                Status          = "$($status.completed) ($($status.downloaded)) $($status.speed) $($status.eta)"
                                PercentComplete = $percent * 100
                            }
                            Write-Progress @progress_args
    
                            $prev_downloaded_bytes = $downloaded_bytes
                            $Global:update_progress = $false
                        }
                    } while ($count -gt 0)
                }
                finally {
                    if ($timer) { $timer.Stop() }
                    if ($timer_event) { Unregister-Event -SubscriptionId $timer_event.Id }
                    if ($target_stream) { $target_stream.Dispose() }
                    # If file exists and $count is not zero or $null, than script was interrupted by user
                    if ((Test-Path $Filename1) -and $count) { Remove-Item -Path $Filename1 }
                }
            }
            finally {
                if ($response) { $response.Dispose() }
                if ($response_stream) { $response_stream.Dispose() }
            }
        }
    }
    Get-FileFromUrl $URL $Filename1
}
  
$Rename = Start-Job -ScriptBlock { Rename-Item -Path "C:\Support\Updates\$KBx$BIT.tmp" -NewName "C:\Support\Updates\$KBx$BIT.msu"
$Rename | Wait-Job | Receive-Job}
<#

This will install multiple Microsoft Standalone Updates from the specified location silently and without rebooting after each update.
You have the option of rebooting after all of the updates have been installed.  Two logs are crated, an output log from WUSA that has
to be read with Event Viewer (.evtx extension) and a transcript log to give you an idea of what's going on (.log extension).

#>

#Create Transcript Log for Troubleshooting
$VerbosePreference = 'Continue'
$LogPath = Split-Path $MyInvocation.MyCommand.Path
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM.dd.yyyy').log"
Start-Transcript $LogPathName

$Location = "C:\Support\Updates"
$Reboot = "No"
$FileTime = Get-Date -format 'MM.dd.yyyy'
$Updates = (Get-ChildItem $Location | Where-Object { $_.Extension -eq '.msu' } | Sort-Object { $_.LastWriteTime } )
$Qty = $Updates.count

if (!(Test-Path $env:systemroot\SysWOW64\wusa.exe)) {
    $Wus = "$env:systemroot\System32\wusa.exe"
}
else {
    $Wus = "$env:systemroot\SysWOW64\wusa.exe"
}

if (Test-Path C:\Support\Updates\Wusa.evtx) {
    Rename-Item C:\Support\Updates\Wusa.evtx C:\Support\Updates\Wusa.$FileTime.evtx
}

foreach ($Update in $Updates) {
    Write-Information "Starting Update $Qty - `r`n$Update"
    Start-Process -FilePath $Wus -ArgumentList ($Update.FullName, '/quiet', '/norestart', "/log:C:\Support\Updates\Wusa.log") -Wait
    Write-Information "Finished Update $Qty"
    if (Test-Path C:\Support\Updates\Wusa.log) {
        Rename-Item C:\Support\Updates\Wusa.log C:\Support\Updates\Wusa.$FileTime.evtx
    }

    Write-Information '-------------------------------------------------------------------------------------------'

    $Qty = --$Qty

    if ($Qty -eq 0) {
        if ($Reboot -like 'Y*') {
            msg console /server:localhost "Updates completed and system is rebooting."
            Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject3 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess
            Restart-Computer
        }
        else {
            msg console /server:localhost "Updates completed and it is advised to close programs and reboot now to complete. It will automatically reboot within 12 hours"
            Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject3 -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Secret -Attachments $Attachment –DeliveryNotificationOption OnSuccess
            Restart-Computer –delay 43200
        }
    }
}

#Close Transcript Log for Troubleshooting
Stop-Transcript
