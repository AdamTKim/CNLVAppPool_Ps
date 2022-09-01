#SCRIPT MUST BE RUN AS ADMIN TO RESTART APPLICATION POOLS

#Syntax for running script is .\<script_name>.ps1 -perc_cap <int> -directory <string> -days <int> -app_pool <string> -restart_attempts <int> -emails <string>
#Parameters entered on the command line for percentage filled cap (perc_cap), directory to delete files (directory), the app pool name (app_pool), the number of attempts to restart the app pool (restart_attempts) and the emails to send any errors to (emails)
param (
    [Alias('p')][int]$perc_cap = 0, #-p
    [Alias('d')][string]$directory = $null, #-d
    [Alias('n')][int]$days = 0, #-n
    [Alias('a')][string]$app_pool = $null, #-a
    [Alias('r')][int]$restart_attempts = 1, #-r
    [Alias('e')][Parameter(Mandatory=$true)][string]$emails = $null #-e
)

import-module webadministration

#Main function that handles which functions to call and whether to send any emails if any parameters are missing
function Main()
{
    #Get current date
    $current_date = Get-Date
    $current_date_unix = Get-Date -uformat %s

    $date_status_log = "Current Unix Timestamp: $current_date_unix, Current Time: $current_date"
    Generate_Log($date_status_log)

    #Check if any parameters have been provided, if none send and email
    If($perc_cap -eq 0 -and [string]::IsNullOrEmpty($directory) -and $days -eq 0 -and [string]::IsNullOrEmpty($app_pool) -and $restart_attempts -eq 1)
    {
        $email_subject = "Missing Parameters for both DISK SPACE function and APP POOL function"
        $ap_ds_status_email = "Parameters for DISK SPACE include -perc_cap, -directory, -days. Parameters for APP POOL include -app_pool -restart_attempts (default 1)"
        Generate_Email($ap_ds_status_email)
        Generate_Log($email_subject)
    }
    #If any parameters have been provided, check to see if all needed to run are provided
    Else
    {
        #Check if perc_cap, directory or days have been entered
        If($perc_cap -ne 0 -or ![string]::IsNullOrEmpty($directory) -or $days -ne 0)
        {
            #If all params (perc_cap, directory, days) have been entered then run Disk_Space function and flag as run
            If($perc_cap -ne 0 -and ![string]::IsNullOrEmpty($directory) -and $days -ne 0)
            {
                $params_status_log_ds = " -perc_cap <$perc_cap> -directory <$directory> -days <$days>" 
                Disk_Space
            }
            #If any params are missing, then find which ones and send in an email notifying user
            Else
            {
                $email_subject = "Missing Parameters for DISK SPACE function"
                $ds_status_email = "Missing parameter(s) "
                If($perc_cap -eq 0)
                {
                    $params_status_log_ds = " MISSING PARAMETERS FOR DISK SPACE FUNCTION"
                    $ds_status_email = $ds_status_email + "-perc_cap "
                }
                If([string]::IsNullOrEmpty($directory))
                {
                    $params_status_log_ds = " MISSING PARAMETERS FOR DISK SPACE FUNCTION"
                    $ds_status_email = $ds_status_email + "-directory "
                } 
                If($days -eq 0)
                {
                    $params_status_log_ds = " MISSING PARAMETERS FOR DISK SPACE FUNCTION"
                    $ds_status_email = $ds_status_email + "-days"
                }
                Generate_Email($ds_status_email)
            }
        }
        #Check if app_pool or restart_attempts have been providedk
        If(![string]::IsNullOrEmpty($app_pool) -or $restart_attempts -ne 1)
        {
            #If all params (app_pool, restart_attempts) have been entered then run App_Pool function and flag as run
            If(![string]::IsNullOrEmpty($app_pool))
            {
                $params_status_log_ap = " -app_pool <$app_pool> -restart_attempts <$restart_attempts>"
                App_Pool
            }
            #If any params are missing, then find which ones and send in an email notifying user
            Else
            {
                $email_subject = "Missing Parameters for APP POOL function"
                $ap_status_email = "Missing parameter(s) -app_pool"
                $params_status_log_ap = " MISSING PARAMETERS FOR APP POOL FUNCTION"
                Generate_Email($ap_status_email)
            }
        }
        $params_status_log = "Parameters Entered:" + $params_status_log_ds  + $params_status_log_ap
        Generate_Log($params_status_log)
    }
}

#Function that deletes files if conditions are met (older than date, drive usage exceeds max)
function Disk_Space()
{
    #Adjust for "older than" date
    $days = $days * -1
    $date_to_delete = $current_date.AddDays($days)
    
    #Concatenate directory with *.* indentifier to delete files within folder and not the folder itself
    If(!(Test-Path $directory))
    {
        $email_subject = "Directory for DISK SPACE function cannot be found"
        $directory_status = "Directory entered for parameter -directory cannot be found"
        Generate_Email($directory_status)
        Generate_Log($directory_status)
        return
    }
    $directory = $directory + '\*.*'

    #Parse drive letter from directory path
    $drive_letter = Split-Path -Path $directory -Qualifier

    #Get information about disk including total size and free space available
    $disk_size = Get-WmiObject win32_logicaldisk -Computername $env:computername -Filter "DeviceID='$drive_letter'" | Foreach-Object {$_.Size/1GB}
    $disk_freespace = Get-WmiObject win32_logicaldisk -Computername $env:computername -Filter "DeviceID='$drive_letter'" | Foreach-Object {$_.Freespace/1GB}

    #Find total space currently being used on disk
    $used_disk_space = [int]$disk_size - [int]$disk_freespace
    $disk_status_log = [string]([math]::Round($used_disk_space, 2)) + " GB (" + [math]::Round((($used_disk_space / $disk_size) * 100), 2) + "%) out of " + [math]::Round($disk_size, 2) + " GB currently used"
    Generate_Log($disk_status_log)

    #Find max capacity before deleting files specified by user above
    $max_fill_capacity = ([int]$perc_cap / 100) * [int]$disk_size

    #If currently used disk space exceeds specified max used space then delete files in user specified directory
    If($used_disk_space -gt $max_fill_capacity)
    {  
        $num_before_del = (dir $directory | measure).Count
        Get-ChildItem $directory -recurse | Where-Object {$_.LastWriteTime -lt $date_to_delete} | Remove-Item
        $num_files_del = $num_before_del - (dir $directory | measure).Count 
        $file_status_log = "$num_files_del files deleted older than $date_to_delete"
        Generate_Log($file_status_log)
    }
}

#Function that attempts to restart an app pool if it isn't currently running
function App_Pool()
{
    #Check if app_pool exists among setup application pools
    If(Test-Path IIS:\AppPools\$app_pool)
    {
        #If app_pool is already running, log and continue
        If((Get-WebAppPoolState -name $app_pool).Value -eq "Started")
        {
            $ap_status_log = "$app_pool was already started"
            Generate_Log($ap_status_log)
        }
        #If app_pool is not running
        Else
        {
            #Attempt to restart app_pool user specified number of times, break when successful
            For($i=0; $i -lt $restart_attempts; $i++)
            {
                If((Get-WebAppPoolState -name $app_pool).Value -ne "Started")
                {
                    Start-WebAppPool "$app_pool"
                    Start-Sleep -s 15
                }
                Else
                {
                    break
                }
            }
            #Check if app_pool is running, if it is not then send an email to the user specified recipients
            If((Get-WebAppPoolState -name $app_pool).Value -ne "Started")
            {
                $email_subject = "$app_pool could not be restarted"
                $ap_status_log = "$app_pool currently has the status of <" + (Get-WebAppPoolState -name $app_pool).Value + "> after $i restart attempt(s) on $env:computername"
                Generate_Email($ap_status_log)
                Generate_Log($ap_status_log)
            }
            Else
            {
                $ap_status_log = "$app_pool currently has the status of <" + (Get-WebAppPoolState -name $app_pool).Value + "> after $i restart attempt(s) on $env:computername"
                Generate_Log($ap_status_log)
            }
        }
    }
    #If app_pool does not exist among setup application pools, email results to user specified recipients
    Else
    {
        $email_subject = "Cannot find application pool"
        $ap_status_log = "Application pool with name $app_pool cannot be found on $env:computername"
        Generate_Email($ap_status_log)
        Generate_Log($ap_status_log)
    }
}

#Generate email to send containing information about the app_pool
#Emails will be sent to all user specified emails
function Generate_Email($body)
{
    $email_from = "NOREPLY@cityofnorthlasvegas.com"
    $email_to = $emails.replace(' ','').split(';')
    $smtp_server = "cnlv-smtp"
    $email_body = [string]$current_date + ": " + $body
    #Iterate through all of the user supplied emails
    For($j = 0; $j -lt $email_to.Length; $j++)
    {
            Send-MailMessage -From $email_from -To $email_to[$j] -Subject $email_subject -Body $email_body -SMTPServer $smtp_server
    }
}

#Generate log file continaing information and results from the user supplied parameters
function Generate_Log($log_string)
{
    #Assign log directory and log file
    $log_directory = (Get-Item -Path ".\").FullName + '\logs\'
    $log_file = (Get-Item -Path ".\").FullName + '\logs\' + $current_date_unix + '.log'

    #If log file directory doesn't exist, create new one
    If(!(Test-Path $log_directory))
    {
        New-Item -ItemType Directory -Force -Path $log_directory
    }
    $current_date_log = Get-Date
    $log_string = [string]$current_date_log + ": " + $log_string
    Add-content $log_file -value $log_string
}

#Execute main function
Main