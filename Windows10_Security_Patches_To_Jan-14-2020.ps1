
# Powershell script to patch Curveball and all previously patched vulnerabilities on Windows 10 1607-1909 
# as this will identify the OS & Build then download the matching patches released up to Jan 14th 2020

# Creating workspace in C:\Support for logs and files
if (![System.IO.Directory]::Exists("C:\support\")) {
    [system.io.directory]::CreateDirectory("c:\Support")
    [system.io.directory]::CreateDirectory("c:\Support\Updates")
}

# Check if windows 10 before continuing to patch
Set-Variable -Name "WINVS" -Value ([Environment]::OSVersion.Version).Major

#Matching OS Build to a KB Identifier & Downloads
switch ((Get-CimInstance Win32_OperatingSystem).BuildNumber) {
    #Windows 7 & 8 
    7600 { $Script:OS = "W2K8R2" }
    7601 { $Script:OS = "W2K8R2SP1" }    
    9200 { $Script:OS = "W2K12" }
    9600 { $Script:OS = "W2K12R2" }
    #Win 10 versions
    10240 { $Script:OS = "W2K10v1507" }
    10586 { $Script:OS = "W2K10v1511" }
    14393 { $Script:OS = "W2K10v1607" }
    14393 { $Script:KB = "KB4534271" }
    14393 { $Script:KB64 = "kb4534271-x64_a009e866038836e277b167c85c58bbf1e0cc5dc8.msu" }
    14393 { $Script:KBx86 = "kb4534271-x86_1401cdaf3781a6b032b558afd90fff6faa5569d3.msu" }
    15063 { $Script:OS = "W2K10v1703" }
    15063 { $Script:KB64 = "W2K10v1703" }
    15063 { $Script:KB86 = "W2K10v1703" }
    16299 { $Script:OS = "W2K10v1709" }
    16299 { $Script:KB64 = "W2K10v1709" }
    16299 { $Script:KB86 = "W2K10v1709" }
    17134 { $Script:OS = "W2K10v1803" }
    17134 { $Script:KB64 = "W2K10v1803" }
    17134 { $Script:KB86 = "W2K10v1803" }
    17763 { $Script:OS = "W2K10v1809" }
    17763 { $Script:KB64 = "W2K10v1809" }
    17763 { $Script:KB86 = "W2K10v1809" }
    18362 { $Script:OS = "W2K10v1903" }
    18362 { $Script:KB = "KB4528760" }
    18362 { $Script:1903 = "W2K10v1903" }
    18362 { $Script:1903x86 = "W2K10v1903" }
    18363 { $Script:OS = "W2K10v1909" }
    18363 { $Script:KB = "KB4528760" } 
    18363 { $Script:1909 = "W2K10v1909" }
    18363 { $Script:1909x86 = "W2K10v1909" }
    19041 { $Script:OS = "W2K10v2004" }
    19041 { $Script:2004 = "W2K10v2004" }
    19041 { $Script:2004x86 = "W2K10v2004" }
    # PENDING: Server 2016
    default { $Script:OS = "Not Listed" }
}

# Logging OS Version info
$OS | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt
Get-HotFix | Select HotFixID, InstalledOn | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append

# Exit if OS is not Win10
If (!($WINVS -eq 10)) {
    #Pending further email info
    Send-MailMessage -from "ServiceEmail <PatcherEmail@Yourdomain.com>"
    -to "Yourname <youreemail@domain.com>",
    #"User03 <user03@example.com>" `
    -subject "$env:computername was Win 7 or 8"
    -body "Cumulative Update check began on $Date"
    -Attachment "c:\support\Updates\$env:computername.hotfixes.txt" -smtpServer mail.yourdomain.com
    Exit
}


#Alert user that system is currently updating
msg console /server:localhost "Hello , your computer needs security patches and will need to remain powered on until complete. You can continue using the system but prepare by saving data as there will be a need to reboot later"
Set-Variable -Name "BUILD" -Value ([Environment]::OSVersion.Version).Build

#Shorten website
$Script:Web = "http://download.windowsupdate.com/d/msdownload/update/software/secu/2020/01/windows10.0-"



#Determine download based on OS Build & "Bitness"
switch ( $KB ) {
    
    W2K10v1507 {"email that this does not have patching" }
    W2K10v1511 {"email that this does not have patching" }
    
    W2K10v1607 { if ((gwmi win32_operatingsystem | select osarchitecture).osarchitecture -eq "64-bit") {
            #64 bit logic here
            Write "64-bit OS Detected" | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append
            $Script:LINK = "LINK" -Value "$Web$KB64"
        }
        else {
            #32 bit logic here
            Write "32-bit OS Detected" | Out-file -Filepath c:\support\Updates\$env:computername.hotfixes.txt -Append
            Set-Variable -Name "LINK" -Value "$Web$KB86"
        }}
  

        
        default { $Script:LINK = "LINK" -Value "ERROR"}
    }
    

            
            
            
            
            
            
    If 1903 or 1909 and have KB4528760, skip
            



 
    if (Test-Path "C:\Support\Updates" -Pathtype Container) {
        Set-Variable -Name "URL" -Value "$LINK"
        Set-Variable -Name "Filename1" -Value "C:\Support\Updates\CumulativeUpdate.tmp"
            
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
    
        $Rename = Start-Job -ScriptBlock { Rename-Item -Path "C:\Support\Updates\CumulativeUpdate.tmp" -NewName "C:\Support\Updates\$OS_$KB_CumulativeUpdate.msu"
        $Rename | Wait-Job | Receive-Job
 
             
        }
    }    
