
# Get-ScheduledTask.ps1
# Written by Bill Stewart (bstewart@iname.com)

#requires -version 2

# Version history:
#
# Version 1.2 (23 May 2012)
# * LastTaskResult property was incorrectly outputting previous task's result
#   when null. Fixed.
#
# Version 1.1 (01 May 2012)
# * Added -Hidden parameter to output tasks marked as hidden.
#
# Version 1.0 (12 Jul 2011)
# * Initial version.

<#
.SYNOPSIS
Outputs scheduled task information.

.DESCRIPTION
Outputs scheduled task information. Requires Windows Vista/Server 2008 or later.

.PARAMETER TaskName
The name of a scheduled task to output. Wildcards are supported. The default value is * (i.e., output all tasks).

.PARAMETER ComputerName
A computer or list of computers on which to output scheduled tasks.

.PARAMETER Subfolders
Specifies whether to support task subfolders (Windows Vista/Server 2008 or later only).

.PARAMETER Hidden
Specifies whether to output hidden tasks.

.PARAMETER ConnectionCredential
The connection to the task scheduler service will be made using these credentials. If you don't specify this parameter, the currently logged on user's credentials are assumed. This parameter only supports connecting to the scheduler service on remote computers running Windows Vista/Server 2008 or later.

.OUTPUTS
PSObjects containing information about scheduled tasks.

.EXAMPLE
PS C:\> Get-ScheduledTask
This command outputs the scheduled tasks in the root tasks folder on the current computer.

.EXAMPLE
PS C:\> Get-ScheduledTask -Subfolders
This command outputs all scheduled tasks on the current computer, including those in subfolders.

.EXAMPLE
PS C:\> Get-ScheduledTask -TaskName \Microsoft\* -Subfolders
This command outputs all scheduled tasks in the \Microsoft task subfolder and its subfolders on the current computer.

.EXAMPLE
PS C:\> Get-ScheduledTask -ComputerName SERVER1
This command outputs scheduled tasks in the root tasks folder on the computer SERVER1.

.EXAMPLE
PS C:\> Get-ScheduledTask -ComputerName SERVER1 -ConnectionCredential (Get-Credential) | Export-CSV Tasks.csv -NoTypeInformation
This command prompts for credentials to connect to SERVER1 and exports the scheduled tasks in the computer's root tasks folder to the file Tasks.csv.

.EXAMPLE
PS C:\> Get-Content Computers.txt | Get-ScheduledTask
This command outputs all scheduled tasks for each computer listed in the file Computers.txt.
#>

[CmdletBinding()]
param(
  [parameter(Position=0)] [String[]] $TaskName="*",
  [parameter(Position=1,ValueFromPipeline=$TRUE)] [String[]] $ComputerName=$ENV:COMPUTERNAME,
  [switch] $Subfolders,
  [switch] $Hidden,
  [System.Management.Automation.PSCredential] $ConnectionCredential
)

begin {
$ErrorActionPreference = SilentlyContinue
  $PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("ComputerName")) -and (-not $ComputerName)
  $MIN_SCHEDULER_VERSION = "1.2"
  $TASK_ENUM_HIDDEN = 1
  $TASK_STATE = @{0 = "Unknown"; 1 = "Disabled"; 2 = "Queued"; 3 = "Ready"; 4 = "Running"}
  $ACTION_TYPE = @{0 = "Execute"; 5 = "COMhandler"; 6 = "Email"; 7 = "ShowMessage"}

  # Try to create the TaskService object on the local computer; throw an error on failure
  try {
    $TaskService = new-object -comobject "Schedule.Service"
  }
  catch [System.Management.Automation.PSArgumentException] {
    throw $_
  }

  # Returns the specified PSCredential object's password as a plain-text string
  function get-plaintextpwd($credential) {
    $credential.GetNetworkCredential().Password
  }

  # Returns a version number as a string (x.y); e.g. 65537 (10001 hex) returns "1.1"
  function convertto-versionstr([Int] $version) {
    $major = [Math]::Truncate($version / [Math]::Pow(2, 0x10)) -band 0xFFFF
    $minor = $version -band 0xFFFF
    "$($major).$($minor)"
  }

  # Returns a string "x.y" as a version number; e.g., "1.3" returns 65539 (10003 hex)
  function convertto-versionint([String] $version) {
    $parts = $version.Split(".")
    $major = [Int] $parts[0] * [Math]::Pow(2, 0x10)
    $major -bor [Int] $parts[1]
  }

  # Returns a list of all tasks starting at the specified task folder
  function get-task($taskFolder) {
    $tasks = $taskFolder.GetTasks($Hidden.IsPresent -as [Int])
    $tasks | foreach-object { $_ }
    if ($SubFolders) {
      try {
        $taskFolders = $taskFolder.GetFolders(0)
        $taskFolders | foreach-object { get-task $_ $TRUE }
      }
      catch [System.Management.Automation.MethodInvocationException] {
      }
    }
  }

  # Returns a date if greater than 12/30/1899 00:00; otherwise, returns nothing
  function get-OLEdate($date) {
    if ($date -gt [DateTime] "12/30/1899") { $date }
  }

  function get-scheduledtask2($computerName) {
    # Assume $NULL for the schedule service connection parameters unless -ConnectionCredential used
    $userName = $domainName = $connectPwd = $NULL
    if ($ConnectionCredential) {
      # Get user name, domain name, and plain-text copy of password from PSCredential object
      $userName = $ConnectionCredential.UserName.Split("\")[1]
      $domainName = $ConnectionCredential.UserName.Split("\")[0]
      $connectPwd = get-plaintextpwd $ConnectionCredential
    }
    try {
      $TaskService.Connect($ComputerName, $userName, $domainName, $connectPwd)
    }
    catch [System.Management.Automation.MethodInvocationException] {
      write-warning "$computerName - $_"
      return
    }
    $serviceVersion = convertto-versionstr $TaskService.HighestVersion
    $vistaOrNewer = (convertto-versionint $serviceVersion) -ge (convertto-versionint $MIN_SCHEDULER_VERSION)
    $rootFolder = $TaskService.GetFolder("\")
    $taskList = get-task $rootFolder
    if (-not $taskList) { return }
    foreach ($task in $taskList) {
      foreach ($name in $TaskName) {
        # Assume root tasks folder (\) if task folders supported
        if ($vistaOrNewer) {
          if (-not $name.Contains("\")) { $name = "\$name" }
        }
        if ($task.Path -notlike $name) { continue }
        $taskDefinition = $task.Definition
        $actionCount = 0
        foreach ($action in $taskDefinition.Actions) {
          $actionCount += 1
          $output = new-object PSObject
          # PROPERTY: ComputerName
          $output | add-member NoteProperty ComputerName $computerName
          # PROPERTY: ServiceVersion
          $output | add-member NoteProperty ServiceVersion $serviceVersion
          # PROPERTY: TaskName
          if ($vistaOrNewer) {
            $output | add-member NoteProperty TaskName $task.Path
          } else {
            $output | add-member NoteProperty TaskName $task.Name
          }
          #PROPERTY: Enabled
          $output | add-member NoteProperty Enabled ([Boolean] $task.Enabled)
          # PROPERTY: ActionNumber
          $output | add-member NoteProperty ActionNumber $actionCount
          # PROPERTIES: ActionType and Action
          # Old platforms return null for the Type property
          if ((-not $action.Type) -or ($action.Type -eq 0)) {
            $output | add-member NoteProperty ActionType $ACTION_TYPE[0]
            $output | add-member NoteProperty Action "$($action.Path) $($action.Arguments)"
          } else {
            $output | add-member NoteProperty ActionType $ACTION_TYPE[$action.Type]
            $output | add-member NoteProperty Action $NULL
          }
          # PROPERTY: LastRunTime
          $output | add-member NoteProperty LastRunTime (get-OLEdate $task.LastRunTime)
          # PROPERTY: LastResult
          if ($task.LastTaskResult) {
            # If negative, convert to DWORD (UInt32)
            if ($task.LastTaskResult -lt 0) {
              $lastTaskResult = "0x{0:X}" -f [UInt32] ($task.LastTaskResult + [Math]::Pow(2, 32))
            } else {
              $lastTaskResult = "0x{0:X}" -f $task.LastTaskResult
            }
          } else {
            $lastTaskResult = $NULL  # fix bug in v1.0-1.1 (should output $NULL)
          }
          $output | add-member NoteProperty LastResult $lastTaskResult
          # PROPERTY: NextRunTime
          $output | add-member NoteProperty NextRunTime (get-OLEdate $task.NextRunTime)
          # PROPERTY: State
          if ($task.State) {
            $taskState = $TASK_STATE[$task.State]
          }
          $output | add-member NoteProperty State $taskState
          $regInfo = $taskDefinition.RegistrationInfo
          # PROPERTY: Author
          $output | add-member NoteProperty Author $regInfo.Author
          # The RegistrationInfo object's Date property, if set, is a string
          if ($regInfo.Date) {
            $creationDate = [DateTime]::Parse($regInfo.Date)
          }
          $output | add-member NoteProperty Created $creationDate
          # PROPERTY: RunAs
          $principal = $taskDefinition.Principal
          $output | add-member NoteProperty RunAs $principal.UserId
          # PROPERTY: Elevated
          if ($vistaOrNewer) {
            if ($principal.RunLevel -eq 1) { $elevated = $TRUE } else { $elevated = $FALSE }
          }
          $output | add-member NoteProperty Elevated $elevated
          # Output the object
          $output
        }
      }
    }
  }
}

process {
  if ($PIPELINEINPUT) {
    get-scheduledtask2 $_
  }
  else {
    $ComputerName | foreach-object {
      get-scheduledtask2 $_
    }
  }
}
