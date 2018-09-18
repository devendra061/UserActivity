########################
#Date 9/12/2108
#
#This program captures the current process and window title. Then, it writes to the log file.
#It creates a log file each day with the file format YYYY-MM-dd-appactivity.log
#It keeps track of the amount of time spent on a active application window.
#It logs the activities in the CSV file format. Log file is located at $path/$logfileName
#It verifies if the script is already running. If it is, then it doesn't run again.
#
#Code insipired from https://social.technet.microsoft.com/Forums/en-US/4d257c80-557a-4625-aad3-f2aac6e9a1bd/get-active-window-info?forum=winserverpowershell
#
#How to Run:
#Use the following command line
#powershell -windowstyle hidden -file AppActivity.ps1
#
#######################
$code = @'
    [DllImport("user32.dll")]
     public static extern IntPtr GetForegroundWindow();
'@
Add-Type $code -Name Utils -Namespace Win32

#Check if this script is already running
$isRunning=@(Get-WmiObject Win32_Process -Filter "Name='powershell.exe' AND CommandLine LIKE '%AppActivity.ps1%'").Count -gt 1
#If this script is running already, don't run it again
if($isRunning) {
	Exit
}



#Change the following log path as you desire; Make sure it trails with backslash (\).
$path = "C:\ISD\Logs\"

#Check if $path exists; If not create the hidden folder
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
	  $f=get-item $path -Force
	  $f.attributes="Hidden"
}

#Uncomment below during codig and testing so that you have a fresh log file to start with
#Remove-Item -path $logfile

#Define previousProcess variable
$previousProcess=""
$previousProcessTitle=""

#Define timeSpent variable; it will keep track of how long the window remained active
$timeSpent=0

#Loop every 5 seconds
while(1){
    $hwnd = [Win32.Utils]::GetForegroundWindow()

#Get the current active window information    
	$process=Get-Process | 
        Where-Object { $_.mainWindowHandle -eq $hwnd }

#If the current active processTitle is same as the previous active processTitle, don't write it to CSV file	
#PreviousProcess will be recorded in CSV file along with the amount of timeSpent 
	If ($previousProcessTitle -ne $process.MainWindowTitle) {

		#Get the current time
			$time=	Get-Date

		#Get the current user	
			$currentUser = $(Get-WMIObject -class Win32_ComputerSystem | select username).username

		#Get the hostname
			$hostname = $(Get-WMIObject -class Win32_ComputerSystem).Name
			
		#Create CustomObjec to create CSV file	
			$allData=[PSCustomObject]@{
				Hostname=$hostname
				CurrentUser=$currentUser
				#Deduct the amount of timeSpent from the current time to get the time user switched to the application
				Time=$time.AddSeconds(-$timeSpent)
				TimeSpent=$timeSpent
				Name=$previousProcess
				Title=$previousProcessTitle
			}


		#Create a log file name with the format of YYYY-MM-dd-appactivity.log
		#This allows to have a daily log
		$logfileName=(Get-Date -Format "yyyy-MM-dd")+"-appactivity.log"
		$logfile=$path+$logfileName

		#Export data to a CSV file
		$allData|Export-Csv -Append -Path $logfile
		
		#Reset timeSpent timer
		$timeSpent=0
	}
	
	$previousProcess=$process.processName
	$previousProcessTitle=$process.MainWindowTitle
	
	#Add 5 seconds to the timeSpent because the Poll Interval is 5 seconds
	$timeSpent=$timeSpent+5
	
#Setup Poll interval of 5 seconds	
    sleep -Milliseconds 5000
}
