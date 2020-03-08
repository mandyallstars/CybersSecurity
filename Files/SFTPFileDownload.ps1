# **************************************************************************************************************
# **************************************************************************************************************
# ** Name:     copy_winpak_files.ps1
# ** Purpose:  1. Remove all exisitng WP CSV files in the local directory
# **           2. Run WinScp to download file from SFTP server Workday Win-Pak inbound folder to the local directory and the archive directory
# **           3. Remove the WP files from remote directory
# **           4. Email concerned parties the result of the file copy operation
# ** Author:   M
# ** Revision: 2019/10/17 - initial version
# **           2019/??/?? - added ...
# **************************************************************************************************************
# **************************************************************************************************************

#set path for the powershell script
Set-Location -Path "C:\WD";

$parent = Get-Location;

#convert the runtime at the time of execution of script to string
$runtime = (Get-Date).toString("yyyyMMdd_HH.mm.ss");

$logsDir = "C:\WD\logs"
$winpakFileArchive = "C:\WD\archiveFolder\$runtime"
$currentWinpakFiles = "C:\WD\WP*.csv"

#create logs directory if it does not exist
If(!(test-path $logsDir))
{
      New-Item -ItemType Directory -Force -Path $logsDir
}

#create a new directory to have an archive based on dates
New-Item -Path "C:\folder\archiveFolder\" -Name "$runtime" -ItemType "directory"

#delete any previous Win Pak files
Remove-Item -Path $currentWinpakFiles

#set log file location
$logfile = "$logsDir\ScriptName_$runtime.log"
$winscplog = "$logsDir\ScriptNameWinSCP_$runtime.log"

#log function to log information as required
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

#set variables for use in script
$FromEmail = 'wd@example.com'
Write-Log INFO "From Email: $FromEmail" $logfile

$Recipient = "<md@example.com>"
Write-Log INFO "Recipient Emails: $Recipient" $logfile

$Subject = "WP CSV File Copy: Success"
$EmailBody = ""

#setting email server information and writing it to LOG file
$PSEmailServer = 'smtp.website.com'
Write-Log INFO "Email Server: $PSEmailServer" $logfile

$Port = 25
Write-Log INFO "Port: $Port" $logfile


#setting variables for remote files
Write-Log INFO "Setting variables for remote files" $logfile
$remotePath = "/wd/wp/inbound/"

try
{
    # Load WinSCP .NET assembly
    Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
    Write-Log INFO "WinSCP .NET assembly loaded" $logfile

    # Setup session options
    Write-Log INFO "Setting WinSCP Session options" $logfile
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = "sftpserver.wesbite.com"
        UserName = "wd"
        SshPrivateKeyPath = "c:\folder\keyFile.ppk"
        SshHostKeyFingerprint = ""
    }
    Write-Log INFO "WinSCP Session options set successfully" $logfile

    #Creating new WinSCP session object
    $session = New-Object WinSCP.Session

    #setting WinSCP log file location
    $session.SessionLogPath = $winscplog
    Write-Log INFO "WinSCP Log file location set to $winscplog" $logfile
    Write-Log INFO "Look in the WinSCP log file for detailed information related to WinSCP Session" $logfile


        #Initiating WinSCP connection
        Write-Log INFO "Initiating WinSCP session with options $sessionOptions" $logfile
        $session.Open($sessionOptions)
        Write-Log INFO "WinSCP session initiated" $logfile

        #Set file download options
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        Write-Log INFO "Files downloading option set to Binary" $logfile

        $wildcard = "TEST_WP_*"

        Write-Log INFO "Getting a list of files matching wildcard" $logfile
        $fileList = $session.EnumerateRemoteFiles($remotePath, $wildcard, [WinSCP.EnumerationOptions]::None)
        Write-Log INFO "Listing of files matching wildcard obtained" $logfile

        Write-Log INFO "Checking number of files in the listing obtained" $logfile
        if ($fileList.Count -lt 1)
        {
            #no matching file found at the remote location
            Write-Log WARN "WP file not found in the remote source location." $logfile

            $Subject = "WP CSV File Copy: Error"

            $EmailBody = $EmailBody+"WP file not found in the remote source location.
            ~More Info on what to do.
            ~Examine the logs and fix any issues before the next run.`n`n"
        }
        elseif($fileList.Count -gt 1)
        {
            #we only want one file to be at the remote location so that we know which file to copy
            Write-Log WARN "More than 1 WP file found in the remote source location." $logfile

            $Subject = "WP CSV File Copy: Error"

            $EmailBody = $EmailBody+"More than 1 WP file found in the remote source location.
            ~More info about what to do.
            ~More info about what to do
            ~Examine the logs and fix any issues before the next run.`n`n"
        }
            else
            {
                
                #file found at remote location - store the file name in a variable
                Write-Log INFO "Storing File name in a variable" $logfile
                foreach($file in $fileList)
                {
                   $fileName = $file
                }

                $sourceFile = $remotePath+$fileName
                $destFile = "c:\folder\TEST_WP.csv"
                
                #download the file to the archive folder
                Write-Log INFO "Downloading WP File" $logfile
                $transferResult = $session.GetFiles($sourceFile, "$winpakFileArchive\$fileName", $False, $transferOptions)

                #check if download was successful
                if (!$transferResult.IsSuccess)
                {
                    # Print error (but continue with other files)
                    Write-Log ERROR "Error downloading the WP file: $($transferResult.Failures[0].Message)" $logfile

                    $Subject = "WP CSV File Copy: Error"
                    $EmailBody = $EmailBody+"Error Downloading the WP file: $($transferResult.Failures[0].Message).
                    ~More info about what to do.
                    ~Examine the logs and fix any issues before the next run.`n`n"
                }
                else
                {
                    
                    # Print success message
                    Write-Log INFO "SUCCESS: WP file was downloaded successfully." $logfile

                    #copy the file from archive folder to parent folder with generic name
                    Copy-Item "$winpakFileArchive\$fileName" -Destination $destFile

                    $Subject = "WP CSV File Copy: Success"
                    
                    $EmailBody = $EmailBody+"WP file was downloaded successfully.`n`n"

                    $removalResult = $session.RemoveFiles($sourceFile)

                    #check if the remote file was successfully removed
                    if($removalResult.IsSuccess)
                    {
                        Write-Log INFO "Source WP file removed from the remote SFTP server" $logfile
                        $EmailBody = $EmailBody + "Source WP file removed from the remote SFTP server.`n`n"
                    }
                    else
                    {
                        $Subject = "WP CSV File Copy: Exception"
                        Write-Log WARN "Unable to remove source WP file from remote SFTP server. Please remove the file manually before next run." $logfile
                        $EmailBody = $EmailBody + "Unable to remove source WP file from remote SFTP server.
                        ~Please remove the file manually before next run.
                        ~Examine the logs and fix any issues before the next run.`n`n"
                    }

               }

            }

        $EmailBody = $EmailBody + "Log files are located at $logfile and $winscplog`n`n"
        $EmailBody = $EmailBody + "Detailed documentation for this whole process is available at https://website.com/pages/page"
        $EmailBody = $EmailBody + "`n`nAt your service,`nIT Admin"


    exit 0

}
catch
{
    Write-Log WARN "WP files could not be downloaded" $logfile
    Write-Log ERROR "Error: $($_.Exception.Message)" $logfile
    $Subject = "WP CSV File Copy: Error"
    $EmailBody = $EmailBody + "Execution of C:\fo,der\powershellScript.ps1 failed with error message: $($_.Exception.Message).`n
                 ~Copy the WP files manually and import data to WP to avoid missing any employee information.
                 `n`nMore info about what to do.
                 `nLog files are located at $logfile and $winscplog`n`n"
    $EmailBody = $EmailBody + "Detailed documentation for this whole process is available at https://www.website.com/pages/page"
    $EmailBody = $EmailBody + "`n`nAt your service,`nWP Admin"

    Write-Log INFO "Sending Email for results of WP file copy operation" $logfile
    Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody

    exit 1
}
finally
{
        #if email body is empty until this point, send a default message advising to look into the issue.
     if(!$EmailBody)
     {
        $Subject = "WP CSV File Copy: Error"
        $EmailBody = "This is the default body of the email. If you see this in your email, please investigate and fix the PKE file downloading script.`n`n
        Log files are located at $logfile and $winscplog`n`n"
        $EmailBody = $EmailBody + "Detailed documentation for this whole process is available at https://www.website.com/pages/page"
        $EmailBody = $EmailBody + "`n`nAt your service,`nWP Admin"
     }
        
     Write-Log INFO "Sending Email for results of WP file copy operation" $logfile
     Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody
     #Disconnect, clean up
     $session.Dispose()

}
