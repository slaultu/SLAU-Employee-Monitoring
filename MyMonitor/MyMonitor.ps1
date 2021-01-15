#requires -version 3.0

<#

Learn more:
 PowerShell in Depth: An Administrator's Guide (http://www.manning.com/jones2/)
 PowerShell Deep Dives (http://manning.com/hicks/)
 Learn PowerShell in a Month of Lunches (http://manning.com/jones3/)
 Learn PowerShell Toolmaking in a Month of Lunches (http://manning.com/jones4/)
 
  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
#>
$scriptpath = $MyInvocation.MyCommand.Path
$global:dir = Split-Path $scriptpath
Write-Output "Script $dir"
$global:last_screenshot = (Get-Date)
$global:idle_screenshoot_interval = 5 #minutes
$global:loop_intensinty_seconds = 10 #seconds. Decreing will react more quicker to windows change but increase CPU usage.
$global:screenshoot_wait_new_window = 5 #seconds. How many seconds wait before screenshot after windows change.
$global:max_idle_time = 10 #max time (seconds) user can be in idel to start counting as inactive/unproductive time

[Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$global:last_user_idle = 0
$global:total_user_idle = 0



Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32 {

    public static class UserInput {

        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO {
            public uint cbSize;
            public int dwTime;
        }

        public static DateTime LastInput {
            get {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }

        public static TimeSpan IdleTime {
            get {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }

        public static int LastInputTicks {
            get {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@





Function Get-ForegroundWindowProcess {

<#
.Synopsis
Get process for foreground window process
.Description
This command will retrieve the process for the active foreground window, ignoring any process with a main window handle of 0.

It will also ignore Task Switching done with Explorer.
.Example
PS C:\> get-foregroundwindowprocess

Handles  NPM(K)    PM(K)      WS(K) VM(M)   CPU(s)     Id ProcessName                                
-------  ------    -----      ----- -----   ------     -- -----------                                
    538      57   124392     151484   885    34.22   4160 powershell_ise
.Link
Get-Process
#>

[cmdletbinding()]
Param()

Try {
  #test if the custom type has already been added
  [user32] -is [Type] | Out-Null
}
catch {
#type not found so add it
Add-Type -typeDefinition @"

using System;
using System.Runtime.InteropServices;

public class User32
{
[DllImport("user32.dll")]
public static extern IntPtr GetForegroundWindow();
}

"@

}

#get the process for the currently active foreground window as long as it has a value
#greater than 0. A value of 0 typically means a non-interactive window. Also ignore
#any Task Switch windows
Get-Process | 
where {$_.MainWindowHandle -eq ([user32]::GetForegroundWindow()) -AND $_.MainWindowHandle -ne 0 -AND $_.Name -ne 'Explorer' -AND $_.Title -notmatch "Task Switching"}

} #end Get-ForegroundWindowProcess

Function Get-WindowTime {

<#
.Synopsis
Monitor time by active window
.Description
This script will monitor how much time you spend based on how long a given window is active. Monitoring will continue until one of the specified triggers is detected. 

By default monitoring will continue for 1 minute. Use -Minutes to specify a different value. You can also specify a trigger by a specific date and time or by detection of a specific process.
.Parameter Time
Monitoring will continue until this datetime value is met or exceeded.
.Parameter Minutes
The numer of minutes to monitor. This is the default behavior.
.Parameter ProcessName
The name of a process that you would see with Get-Process, e.g. Notepad or Calc. Monitoring will stop when this process is detected.
Parameter AsJob
Run the monitoring in a background job. Note that if you stop the job you will NOT have any results.
.Example
PS C:\> $data = Get-WindowTime -minutes 60

Monitor window activity for the next 60 minutes. Be aware that you won't get your prompt back until this command completes.
.Example
PS C:\> Get-WindowTime -processname calc -asjob

Start monitoring windows in the background until the Calculator process is detected.
.Notes
Last Updated: August 26, 2014
Version     : 1.0

.Link
Get-Process

#>

[cmdletbinding(DefaultParameterSetName= "Minutes")]
Param(
[Parameter(ParameterSetName="Time")]
[ValidateNotNullorEmpty()]
[DateTime]$Time,

[Parameter(ParameterSetName="Minutes")]
[ValidateScript({ $_ -ge 1})]
[Int]$Minutes = 1,

[Parameter(ParameterSetName="Process")]
[ValidateNotNullorEmpty()]
[string]$ProcessName,

[switch]$AsJob

)

Write-Output "[$(Get-Date)] Starting $($MyInvocation.Mycommand)"  

#define a scriptblock to use in the While loop
Switch ($PSCmdlet.ParameterSetName) {

"Time"    {    
            Write-Output "[$(Get-Date)] Stop monitoring at $Time"      
            [scriptblock]$Trigger = [scriptblock]::Create("(get-date) -ge ""$time""")
            Break
          }
"Minutes" {
            $Quit = (Get-Date).AddMinutes($Minutes)
            Write-Output "[$(Get-Date)] Stop monitoring in $minutes minute(s) at $Quit"
            [scriptblock]$Trigger = [scriptblock]::Create("(get-date) -ge ""$Quit""")
            Break  
          }
"Process" {
            Write-Output "[$(Get-Date)] Stop monitoring after trigger $Processname"
           [scriptblock]$Trigger = [scriptblock]::Create("Get-Process -Name $ProcessName -ErrorAction SilentlyContinue")
            Break
          }

} #switch


#define the entire command as a scriptblock so it can be run as a job if necessary
$main = {
    Param($sb)

    if (-Not ($sb -is [scriptblock])) {
     #convert $sb to a scriptblock
     Write-Output "Creating sb from $sb"
     $sb = [scriptblock]::Create("$sb")
    }
    #create a hashtable
    $hash=@{}
    #create a collection of objects
    $objs = @()
    New-Variable -Name LastApp -Value $Null
    while( -Not (&$sb) ) {
        $Process = Get-ForegroundWindowProcess
         [string]$app = $process.MainWindowTitle
         if ( (-Not $app) -AND $process.MainModule.Description ) {
            #if no title but there is a description, use that
            $app = $process.MainModule.Description
         } elseif (-Not $app) {
            #otherwise use the module name
            $app = $process.mainmodule.modulename
         }
        
        
        
          ################################################
          #idle calculation
          ################################################

          $Idle = [math]::Round([PInvoke.Win32.UserInput]::IdleTime.Seconds)
          if ($Idle -gt $global:max_idle_time){
               $global:last_user_idle = $Idle
               Write-Output "[$(Get-Date)] last_user_idle $global:last_user_idle ..."
          } else {
              if ($global:last_user_idle -gt $global:max_idle_time) {
                $global:total_user_idle = $global:total_user_idle + $global:last_user_idle
              }
              $global:last_user_idle = 0
              $diff= (New-TimeSpan -Start $global:last_screenshot -End (Get-Date)).TotalMinutes
              #Write-Output "[$(Get-Date)] Last sreenshoot: $diff"
              if ($diff -gt $global:idle_screenshoot_interval) {
                 Write-Output "[$(Get-Date)] Screenshoot update $global:idle_screenshoot_interval min."
                 screenshot
              }        
          }
          Write-Output "[$(Get-Date)] IDLE: $Idle    LI: $global:last_user_idle     GI: $global:total_user_idle"        
        
        
        
        if ($process -AND (($Process.MainWindowHandle -ne $LastProcess.MainWindowHandle) -OR ($app -ne $lastApp )) ) {
            Write-Output "[$(Get-Date)] NEW App changed to $app"
             #record $last
             if ($LastApp) {
                    #create new object
                    #include a detail property object
                    [pscustomObject]$detail=@{
                      StartTime = $start
                      EndTime = Get-Date
                      ProcessID = $lastProcess.ID
                      Process = if ($LastProcess) {$LastProcess} else {$process}
                    }
                    Write-Output "[$(Get-Date)] 1111 Creating new object for $LastApp"
                    Write-Output "[$(Get-Date)] Time = $([math]::Round($sw.ElapsedMilliseconds/1000))"
                    $obj = New-Object -TypeName PSobject -Property @{
                        WindowTitle = $LastApp
                        Application = $LastMainModule.Description #$LastProcess.MainModule.Description
                        Product = $LastMainModule.Product #$LastProcess.MainModule.Product
                        Time = [math]::Round($sw.ElapsedMilliseconds/1000)
                        Date = (get-date).ToString(‘yyyy-MM-dd’)
                        Idle = [math]::Round($global:total_user_idle)
                        User = $env:UserName
                        Detail = ,([pscustomObject]@{
                         StartTime = $start
                         EndTime = Get-Date
                         ProcessID = $lastProcess.ID
                         Process = if ($LastProcess) {$LastProcess} else {$process}
                        } )
					    StartTime = $start
                        EndTime = Get-Date
                        Process = if ($LastProcess) {$LastProcess} else {$process}                   
					    }    
                    $obj.psobject.TypeNames.Insert(0,"My.Monitored.Window")
                    #add a custom type name
                    #add the object to the collection
                    $objs += $obj
                    $obj | Select-Object -Property User, Date, StartTime, EndTime, Time, Idle, Application, WindowTitle, Product, Title, Process | Export-CSV $global:dir\Data_$env:UserName.csv -Append -NoTypeInformation -Encoding utf8
                    $global:total_user_idle = 0
                    Write-Output "[$(Get-Date)] Sleeping $global:screenshoot_wait_new_window s...."
                    Start-Sleep -Milliseconds ($global:screenshoot_wait_new_window * 1000)
                    screenshot
                    
            } else { #if $lastApp was defined
                Write-Output "[$(Get-Date)] You should only see this once"
            }
            #new Process with a window
            Write-Output "[$(Get-Date)] Start a timer"
            $SW = [System.Diagnostics.Stopwatch]::StartNew()     
            $start = Get-Date
            #set the last app
            $LastApp = $app
            #preserve process information
            $LastProcess = $Process
            $LastMainModule = $process.mainmodule
          #clear app just in case
          Remove-Variable app
      }





           ###sleep script
          Write-Output "[$(Get-Date)] Sleeping $global:loop_intensinty_seconds ..."
          Start-Sleep -Milliseconds ($global:loop_intensinty_seconds * 1000)


    } #while
    #create new object
    Write-Output "[$(Get-Date)] 2222 Creating new object 2222"
    #Write-Output "[$(Get-Date)] Time = $([math]::Round($sw.ElapsedMilliseconds/1000))"

    $obj = New-Object -TypeName PSobject -Property @{
        WindowTitle = $LastApp
        Application = $LastMainModule.Description #$LastProcess.MainModule.Description
        Product = $LastMainModule.Product #$LastProcess.MainModule.Product
        Time = [math]::Round($sw.ElapsedMilliseconds/1000)
        Idle = [math]::Round($global:total_user_idle)
        Date = (get-date).ToString(‘yyyy-MM-dd’)
        User = $env:UserName
        Detail = ,([pscustomObject]@{
         StartTime = $start
         EndTime = Get-Date
         ProcessID = $lastProcess.ID
         Process = if ($LastProcess) {$LastProcess} else {$process}
         })
        }
    
    $obj.psobject.TypeNames.Insert(0,"My.Monitored.Window")
    #add a custom type name
    #add the object to the collection
    $objs += $obj
    $obj | Export-CSV $global:dir\Data_$env:UserName.csv -Append -NoTypeInformation
    $global:total_user_idle = 0
    #} #else create new object
    $objs
    Write-Output "[$(Get-Date)] Ending $($MyInvocation.Mycommand)"  
} #main


if ($asJob) {
    Write-Output "[$(Get-Date)] Running as background job"
   Start-Job -ScriptBlock $main -ArgumentList @($Trigger) -InitializationScript  {Import-Module MyMonitor}

}
else {
    #run it
   Invoke-Command -ScriptBlock $main -ArgumentList @($Trigger)
}

} #end Get-WindowTime

Function Measure-WindowTotal {

<#
.Synopsis
Measure Window usage results.
.Description
This command is designed to take output from Get-WindowTime and measure total time either by Application, the default, by Product or Window title. Or you can elect to get the total time for all piped in measured window objects.

You can also filter based on keywords found in the window title. See examples.
.Parameter Filter
A string to be used for filtering based on the window title. The string can be a regular expression pattern.
.Example
PS C:\> $data = Get-WindowTime -ProcessName calc  
PS C:\> $data | Measure-WindowTotal  

Application                                       TotalTime
-----------                                       ---------
Windows Explorer                                  00:00:00.1487538
Nitro Reader 3                                    00:00:16.2350978
Microsoft PowerPoint                              00:00:22.4399588
Thunderbird                                       00:00:24.5278000
Microsoft Word                                    00:00:27.9303417
Google Chrome                                     00:13:15.7782912
Windows PowerShell                                00:15:50.7087444
Windows PowerShell ISE                            00:24:20.2946657

PS C:\> $data  | Measure-WindowTotal -Product

Product                                           TotalTime
-------                                           ---------
Nitro Reader 3                                    00:00:16.2350978
Thunderbird                                       00:00:24.5278000
Microsoft Office 2013                             00:00:50.3703005
Google Chrome                                     00:13:15.7782912
Microsoft® Windows® Operating System              00:40:11.1521639

The first command gets data from active window usage. The second command measures the results by Application. The last command measures the same data but by the product property.
.Example
PS C:\> $data | Measure-WindowTotal -filter "facebook" -TimeOnly

Days              : 0
Hours             : 0
Minutes           : 5
Seconds           : 1
Milliseconds      : 237
Ticks             : 3012378186
TotalDays         : 0.00348654882638889
TotalHours        : 0.0836771718333333
TotalMinutes      : 5.02063031
TotalSeconds      : 301.2378186
TotalMilliseconds : 301237.8186

Get just the time that was spent on any window that had Facebook in the title.
.Example
PS C:\> $data | Measure-WindowTotal -filter "facebook|twitter"

Application                                       TotalTime
-----------                                       ---------
Google Chrome                                     00:05:55.4317420

Display how much time was spent on Facebook or Twitter.
.Notes
Last Updated: July 3, 2014
Version     : 1.0

.Link
Get-WindowTime
#>

[cmdletbinding(DefaultParameterSetName="Default")]
Param(
[Parameter(Position=0,ValueFromPipeline)]
$InputObject,
[Parameter(ParameterSetName="Product")]
[Switch]$Product,
[Parameter(ParameterSetName="Title")]
[Switch]$Title,
[ValidateNotNullorEmpty()]
[String]$Filter=".*",
[Switch]$TimeOnly
)

Begin {
    Write-Output "[$(Get-Date)] Starting $($MyInvocation.Mycommand)"  

    #initialize
    $hash=@{}

    if ($Product) {
      $objFilter = "Product"
    }
    elseif($Title) {
      $objFilter = "WindowTitle"
    }
    else {
      $objFilter = "Application"
    }

    Write-Output "[$(Get-Date)] Calculating totals by $objFilter"
} #begin

Process {

  #only process objects where the window title matches the filter which
  #by default is everything and there is only one object in the product
  #which should eliminate task switching data
  if ($Inputobject.WindowTitle -match $filter -AND $Inputobject.Product.count -eq 1) {

      if ($hash.ContainsKey($InputObject.$objFilter)) { 
        #update an existing entry in the hash table
         $hash.Item($InputObject.$objFilter) += $InputObject.Time 
      }
      else {
        #Add an entry to the hashtable
        Write-Output "[$(Get-Date)] Adding $($Inputobject.$objFilter)"
        $hash.Add($Inputobject.$objFilter,$Inputobject.time)
      }
  }

} #process

End {
    Write-Output  "[$(Get-Date)] Ending $($MyInvocation.Mycommand)"
    #turn hash table into a custom object and sort on time by default
      $output = $hash.GetEnumerator() | foreach {
      New-Object -TypeName PSObject -Property @{$objfilter=$_.Name;"TotalTime"=$_.Value}
    } | Sort TotalTime
    
    if ($TimeOnly) {
     $output | foreach -begin {$total = New-TimeSpan} -process {$total+=$_.Totaltime} -end {$total}
     }
     else {
     $output | Select $objFilter,TotalTime
     }
} #end

} #end Measure-WindowTotal

Function Get-WindowTimeSummary {
<#
.Synopsis
Get a summary of window usage time
.Description
This command will take an array of window usage data and present a summary based on application. The output will include the total time as well as the first and last times for that particular application.

As an alternative you can get a summary by Product or you can filter using a regular expression pattern on the window title.
.Example
PS C:> 

PS C:\> get-windowtimeSummary $data

Name                          Total                    Start                     End
----                          -----                    -----                     ---
Windows PowerShell            00:03:00.5083791         8/26/2014 8:58:02 AM      8/26/2014 9:44:14 AM
Windows PowerShell ISE        00:32:02.6150152         8/26/2014 8:58:06 AM      8/26/2014 9:58:01 AM
Virtual Machine Connection    00:13:47.6952161         8/26/2014 8:59:03 AM      8/26/2014 9:25:49 AM
Notepad                       00:00:34.3654881         8/26/2014 9:03:14 AM      8/26/2014 9:04:01 AM
SAPIEN PowerShell Studio 2014 00:00:51.5242522         8/26/2014 9:03:42 AM      8/26/2014 9:09:01 AM
Google Chrome                 00:06:34.6744841         8/26/2014 9:23:05 AM      8/26/2014 9:49:45 AM
Thunderbird                   00:01:55.6696204         8/26/2014 9:18:22 AM      8/26/2014 9:19:28 AM
Microsoft Management Console  00:00:37.2989472         8/26/2014 9:19:32 AM      8/26/2014 9:26:01 AM
Microsoft Word                00:00:34.3653455         8/26/2014 9:49:41 AM      8/26/2014 9:49:41 AM
.Example
PS C:\> get-windowtimesummary $data -filter "facebook|hootsuite"

Name                          Total                    Start                     End
                          -----                    -----                         ---
facebook|hootsuite            00:05:27.6352617         8/26/2014 9:23:05 AM      8/26/2014 9:49:45 AM
.Notes
Last Updated: August 26, 2014
Version     : 1.0

.Link
Get-WindowTime
Measure-WindowTotal
#>

[cmdletbinding(DefaultParameterSetName="Type")]
Param(
[Parameter(
Position=0,Mandatory,
HelpMessage="Enter a variable with your Window usage data")]
[ValidateNotNullorEmpty()]
$Data,
[Parameter(ParameterSetName="Type")]
[ValidateSet("Product","Application")]
[string]$Type="Application",

[Parameter(ParameterSetName="Filter")]
[ValidateNotNullorEmpty()]
[string]$Filter

)

Write-Output -Message "[$(Get-Date)] Starting $($MyInvocation.Mycommand)"  


if ($PSCmdlet.ParameterSetName -eq 'Type') {
    Write-Output "Processing on $Type"
    #filter out blanks and objects with multiple products from ALT-Tabbing
    $grouped = $data | Where {$_.$Type -AND $_.$Type.Count -eq 1} | Group-Object -Property $Type
}
else {
  #use filter
  Write-Output "Processing on filter: $Filter"
  $grouped = $data | where {$_.WindowTitle -match $Filter -AND $_.Product.Count -eq 1} |
  Group-Object -Property {$Filter}
}

if ($Grouped) {
    $grouped| Select Name,
    @{Name="Total";Expression={ 
    $_.Group | foreach -begin {$total = New-TimeSpan} -process {$total+=$_.time} -end {$total}
    }},
    @{Name="Start";Expression={
    ($_.group | sort Detail).Detail[0].StartTime
    }},
    @{Name="End";Expression={
    ($_.group | sort Detail).Detail[-1].EndTime
    }}
}
else {
    Write-Warning "No items found"
  }

    Write-Output -Message "Ending $($MyInvocation.Mycommand)"

} #end Get-WindowsTimeSummary

function screenshot() 
{
    $global:last_screenshot = (GET-DATE)
    $DateString = (get-date).ToString(‘yyyy-MM-dd_HHmmss’)
    $path = $global:dir+"\screenshot\"+$env:UserName+"_"+$DateString+".png"
    $width = 0;
    $height = 0;
    $workingAreaX = 0;
    $workingAreaY = 0;

    Write-Output "Taking screenshoot: $path"

    $screen = [System.Windows.Forms.Screen]::AllScreens;

    foreach ($item in $screen)
    {
        if($workingAreaX -gt $item.WorkingArea.X)
        {
            $workingAreaX = $item.WorkingArea.X;
        }

        if($workingAreaY -gt $item.WorkingArea.Y)
        {
            $workingAreaY = $item.WorkingArea.Y;
        }

        $width = $width + $item.Bounds.Width;

        if($item.Bounds.Height -gt $height)
        {
            $height = $item.Bounds.Height;
        }
    }

    $bounds = [Drawing.Rectangle]::FromLTRB($workingAreaX, $workingAreaY, $width, $height); 
    $bmp = New-Object Drawing.Bitmap $width, $height;
    $graphics = [Drawing.Graphics]::FromImage($bmp);
    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size);
	
	#screenshoot quality
	$quality = 75
	
	#screenshoot size in percent
	$scale = 50
	[int32]$new_width = $width * ($Scale / 100)
	[int32]$new_height = $height * ($Scale / 100)
	
	
	$bmp2 = New-Object System.Drawing.Bitmap($new_width, $new_height)
    $graphics2 = [System.Drawing.Graphics]::FromImage($bmp2)
    $graphics2.DrawImage($bmp, 0, 0, $new_width, $new_height)
	
    #Encoder parameter for image quality 
    $myEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1) 
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $quality)
    
    # get codec
    $myImageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|where {$_.MimeType -eq 'image/jpeg'}


    $bmp2.Save($path,$myImageCodecInfo, $($encoderParams));

    $graphics.Dispose();
    $bmp.Dispose();
	$graphics2.Dispose();
    $bmp2.Dispose();
}

#set default display property set
Update-TypeData -TypeName "my.monitored.window" -DefaultDisplayPropertySet "Time","Application","WindowTitle","Product" -DefaultDisplayProperty WindowTitle -Force
Update-TypeData -TypeName "deserialized.my.monitored.window" -DefaultDisplayPropertySet "Time","Application","WindowTitle","Product" -DefaultDisplayProperty WindowTitle  -Force

#add an alias for the WindowTitle property
Update-TypeData -TypeName "My.Monitored.Window" -MemberType AliasProperty -MemberName Title -Value WindowTitle -force
Update-TypeData -TypeName "deserialized.My.Monitored.Window" -MemberType AliasProperty -MemberName Title -Value WindowTitle -force

#define some command aliases
Set-Alias -name mwt -Value Measure-WindowTotal
Set-Alias -name gfwp -Value Get-ForegroundWindowProcess
Set-Alias -Name gwt -Value Get-WindowTime
Set-Alias -name gwts -Value Get-WindowTimeSummary

#Get-WindowTime
Get-WindowTime -ProcessName cleanmgr  


#$ScriptToRun= $PSScriptRoot+"\MyMonitor.ps1"
#&$ScriptToRun


#get-wmiobject win32_process | where{ $_.Handle -eq 8 } |
#Select Name, @{Name="UserName";Expression={$_.GetOwner().Domain+"\"+$_.GetOwner().User}} | 
#Sort-Object UserName, Name