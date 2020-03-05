# **************************************************************************************************************
# **************************************************************************************************************
# ** Name:     update_sql_table_winpak.ps1
# ** Purpose:  1. Use the file imported by Powershell script copy_winpak_files.ps1 - TEST_Win-Pak.csv
# **           2. Import the data from TEST_Win-Pak.csv file into SQL seerver Win-Pak database to update dbo.TestCardHolder table
# **           4. Email concerned parties the result of the file copy operation
# ** Author:   Mandeep Dhillon
# ** Revision: 2019/12/03 - initial version
# **           2019/??/?? - added ...
# **************************************************************************************************************
# **************************************************************************************************************

#set path for the powershell script
Set-Location -Path "C:\workday";

$parent = Get-Location;

cd $parent

#convert the runtime at the time of execution of script to string
$runtime = (Get-Date).toString("yyyyMMdd_HH.mm.ss");

$logsDir = "C:\workday\logs"

#create logs directory if it does not exist
If(!(test-path $logsDir))
{
      New-Item -ItemType Directory -Force -Path $logsDir
}

#set log file location
$logfile = "$logsDir\update_sql_table_winpak_$runtime.log"


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

$Subject = "WP EE Data Import: Success"
$EmailBody = ""

#setting email server information and writing it to LOG file
$PSEmailServer = 'smtp.example.com'
Write-Log INFO "Email Server: $PSEmailServer" $logfile

$Port = 25
Write-Log INFO "Port: $Port" $logfile


#Setting constants for SQL server information
$serverInstance = "ServerName"
$database = "Database Name"
$schema = "dbo"
$cardHolderTable = "TableName"
$cardTable = "TableNameSecond"
$columname = "RecordID"

#setting CSV import file name
$importFileName = "$parent\TEST_WP.csv"


#importing module for SQL server
Write-Log INFO "Importing SQL Server Module" $logfile
Import-Module SqlServer -ErrorAction SilentlyContinue

#if the previous statement was unsuccessful, output the errors to
#logfile, send an email and exit program
if(!$?) {

    Write-Log ERROR "Could not import SQL Server Module" $logfile
    Write-Log ERROR "ExceptionMessage: $($error[0].Exception.Message)" $logfile
    Write-Log ERROR "Target Object: $($error[0].TargetObject)" $logfile
    Write-Log ERROR "Category Info: $($error[0].CategoryInfo)" $logfile
    Write-Log ERROR "ErrorID: $($error[0].FullyQualifiedErrorId)" $logfile

    $Subject = "WP EE Data Import: Error"

    $EmailBody = $EmailBody = $EmailBody + "Execution of C:\folder\powershellScript.ps1 failed with error message: $($error[0].Exception.Message).`n
            `nMore Error.
            ~More Error.
            ~More Error.`n`n"


    $EmailBody = $EmailBody + "Log file is located at $logfile"

    $EmailBody = $EmailBody + "`n`nDetailed documentation for this whole process is available at https://www.website.com/"
    $EmailBody = $EmailBody + "`n`nAt your service,`nIT Admin"
    
    Write-Log INFO "Sending Email for results of WP EE Data Import operation" $logfile
    Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody

    exit 1

} else {

    Write-Log INFO "SQL Server Module Imported Successfully" $logfile

    $EmailBody = $EmailBody + "SQL Server Module Imported Successfully`n"
}


#setting the database
Write-Log INFO "Getting WP database" $logfile
Get-SqlDatabase -ServerInstance "WPServer" -ErrorAction SilentlyContinue


#if the previous statement was unsuccessful, output the errors to
#logfile, send an email and exit program
if(!$?) {

    Write-Log ERROR "Could not get the requested SQL Server database - WP" $logfile
    Write-Log ERROR "ExceptionMessage: $($error[0].Exception.Message)" $logfile
    Write-Log ERROR "Target Object: $($error[0].TargetObject)" $logfile
    Write-Log ERROR "Category Info: $($error[0].CategoryInfo)" $logfile
    Write-Log ERROR "ErrorID: $($error[0].FullyQualifiedErrorId)" $logfile

    $Subject = "WP EE Data Import: Error"

    $EmailBody = $EmailBody + "Execution of C:\folder\powershellScript.ps1 failed with error message: $($error[0].Exception.Message).`n
            `nCould not get the requested SQL Server database - WP.
            ~Additional error information can be found in the Log file.
            ~Examine the logs and fix any issues before the next run.`n`n"

    $EmailBody = $EmailBody + "Log file is located at $logfile"

    $EmailBody = $EmailBody + "`n`nDetailed documentation for this whole process is available at https://www.website.com/"
    $EmailBody = $EmailBody + "`n`nAt your service,`nIT Admin"
    
    Write-Log INFO "Sending Email for results of WP EE Data Import operation" $logfile
    Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody

    exit 1

} else {

    Write-Log INFO "Got a SQL Server database object for 'WP'" $logfile

    $EmailBody = $EmailBody + "Got a SQL Server database object for 'WP'`n"

}

#importing CSV file into a variable
Write-Log INFO "Importing the CSV data file $importFileName into a variable" $logfile
$employees = Import-Csv $importFileName -ErrorAction SilentlyContinue

#if the previous statement was unsuccessful, output the errors to
#logfile, send an email and exit program
if(!$?) {

    Write-Log ERROR "Could not import the CSV data file $importFileName" $logfile
    Write-Log ERROR "ExceptionMessage: $($error[0].Exception.Message)" $logfile
    Write-Log ERROR "Target Object: $($error[0].TargetObject)" $logfile
    Write-Log ERROR "Category Info: $($error[0].CategoryInfo)" $logfile
    Write-Log ERROR "ErrorID: $($error[0].FullyQualifiedErrorId)" $logfile

    $Subject = "WP Employee Data Import: Error"

    $EmailBody = $EmailBody + "Execution of $parent\powershellScript.ps1 failed with error message: $($error[0].Exception.Message).`n
            `nCould not import the CSV data file $importFileName.
            ~Additional error information can be found in the Log file.
            ~Examine the logs and fix any issues before the next run.`n`n"

    $EmailBody = $EmailBody + "Log file is located at $logfile"

    $EmailBody = $EmailBody + "`n`nDetailed documentation for this whole process is available at https://www.website.com/"
    $EmailBody = $EmailBody + "`n`nAt your service,`nIT Admin"
    
    Write-Log INFO "Sending Email for results of WP Employee Data Import operation" $logfile
    Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody

    exit 1

} else {

    Write-Log INFO "CSV data file $importFileName imported successfully" $logfile

    $EmailBody = $EmailBody + "CSV data file $importFileName imported successfully'`n"

}

try  {

foreach ($empl in $employees) {
    $emplID = $empl.Employee_ID
    $userName = $empl.User_Name
    $firstName = $empl.First_Name
    $middleName = $empl.Middle_Name
    $lastName = $empl.Last_Name
    $location = $empl.Location
    $jobTitle = $empl.Job_Title
    $department = $empl.Cost_Center
    $managerInfo = $empl.Manager_name_Manager_Position
    $emplStatus = $empl.Status
    $TPHD_flag = $empl.TPHD_flag

    if($middleName){
        $firstMiddleName = $firstName + ' ' + $middleName
    } else {
        $firstMiddleName = $firstName
    }

    $dbCmdUpdateUser = @"    
    USE [$database]
    SET QUOTED_IDENTIFIER OFF
    UPDATE [$schema].[$cardHolderTable]
    SET FirstName = "$firstMiddleName"
    , LastName = "$lastName"
    , Note11 = "$location"
    , Note12 = "$jobTitle"
    , Note13 = "$department"
    , Note14 = "$managerInfo"
    , Note15 = "$emplStatus"
    , Note16 = "$TPHD_flag"
    WHERE Note10 = $emplID;
    SET QUOTED_IDENTIFIER ON
"@

#running SQLUpdate command on the database
Write-Log INFO "Updating data for Employee $emplID" $logfile
Invoke-SqlCmd -ServerInstance $serverInstance -Query $dbCmdUpdateUser  -ErrorAction SilentlyContinue

#if the previous statement was unsuccessful, output the errors to
#logfile and continue on until the loop finishes
    if(!$?) {

        Write-Log ERROR "There was an error updating $emplID data in WP database" $logfile
        Write-Log ERROR "ExceptionMessage: $($error[0].Exception.Message)" $logfile
        Write-Log ERROR "Target Object: $($error[0].TargetObject)" $logfile
        Write-Log ERROR "Category Info: $($error[0].CategoryInfo)" $logfile
        Write-Log ERROR "ErrorID: $($error[0].FullyQualifiedErrorId)" $logfile

        $Subject = "WP EE Data Import: Error"

        $EmailBody = $EmailBody+"$emplID data update failed`n"

    } else {

        Write-Log INFO "Data for EE $emplID updated successfuly" $logfile

        $EmailBody = $EmailBody + "$emplID data updated successfully`n"

    }

    if($emplStatus -eq "Terminated") {

        #retrieving RecordID for the Terminated employee
        $userRecord = Read-SqlTableData -ServerInstance $serverInstance -DatabaseName $database -SchemaName $schema -TableName $cardHolderTable | Where {$_.Note10 -eq $emplID}

        $recordID = $($userRecord.RecordID)

        Write-Log INFO "$emplID status is Terminated. Running Process to deactivate their access cards" $logfile

        $dbCmdUpdateCard = @"    
            USE [$database]
            SET QUOTED_IDENTIFIER OFF
            UPDATE [$schema].[$cardTable]
            SET CardStatus = 2
            WHERE CardHolderID = $recordID
            and Deleted != 1;
            SET QUOTED_IDENTIFIER ON
"@

        #running SQLUpdate command on the database
        Write-Log INFO "Updating Card Status for Terminated Employee $emplID" $logfile
        Invoke-SqlCmd -ServerInstance $serverInstance -Query $dbCmdUpdateCard  -ErrorAction SilentlyContinue

        #if the previous statement was unsuccessful, output the errors to
        #logfile and continue on until the loop finishes
        if(!$?) {

            Write-Log ERROR "There was an error updating Terminated $emplID card status in WP database" $logfile
            Write-Log ERROR "ExceptionMessage: $($error[0].Exception.Message)" $logfile
            Write-Log ERROR "Target Object: $($error[0].TargetObject)" $logfile
            Write-Log ERROR "Category Info: $($error[0].CategoryInfo)" $logfile
            Write-Log ERROR "ErrorID: $($error[0].FullyQualifiedErrorId)" $logfile

            $Subject = "WP Employee Data Import: Error"

            $EmailBody = $EmailBody+"Terminated employee $emplID Card Status update failed`n"

        } else {
        
            Write-Log INFO "Card Status for Terminated Employee $emplID updated successfuly" $logfile

            $EmailBody = $EmailBody + "Terminated employee $emplID Card Status updated successfully`n"

        }
    
    }


}


} catch {

    Write-Log WARN "There was an error in updating data in WP database" $logfile
    Write-Log ERROR "Error: $($_.Exception.Message)" $logfile

    $Subject = "WP Employee Data Import: Error"

    $EmailBody = $EmailBody + "Execution of C:\folder\powershellScript.ps1 failed with error message: $($_.Exception.Message).`n
                 `n`nExamine the logs and fix any issues before the next run.
                 `nLog file is located at $logfile`n`n"

    $EmailBody = $EmailBody + "Detailed documentation for this whole process is available at https://www.website.com/"

    $EmailBody = $EmailBody + "`n`nAt your service,`nWP Admin"

    Write-Log INFO "Sending Email for results of WP Employee Data Import operation" $logfile
    Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody

    exit 1


} finally {

   #if email body is empty until this point, send a default message advising to look into the issue.
     if(!$EmailBody)
     {
        $Subject = "WP Employee Data Import: Error"

        $EmailBody = "Execution of C:\folder\powershellScript.ps1 failed.`n
        This is the default body of the email. If you see this in your email, please investigate and fix the WINPAK data import script`n`n
        Log files are located at $logfile`n`n"

        $EmailBody = $EmailBody + "Detailed documentation for this whole process is available at https://www.website.com/"
        $EmailBody = $EmailBody + "`n`nAt your service,`nWP Admin"
     } elseif($Subject -like "*Success") {

        $EmailBody = "Execution of C:\folder\powershellScript.ps1 was successful`n
                     ~Additional information can be found in log file which is located at $logfile`n
                     ~Detailed documentation for this whole process is available at https://www.website.com/
                     `nSee below for additional details.
                     `nAt your service,`nWin-Pak Admin`n`n"+$EmailBody
        
     }
        
     Write-Log INFO "Sending Email for results of WP Employee Data Import operation" $logfile
     Send-MailMessage -From $FromEmail -To $Recipient -Subject $Subject -Body $EmailBody
     #Disconnect, clean up
     exit 0

}