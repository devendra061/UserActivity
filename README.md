# UserActivity

#This program captures the current process and window title. Then, it writes to the log file.

#It creates a log file each day with the file format YYYY-MM-dd-appactivity.log

#It keeps track of the amount of time spent on a active application window.

#It logs the activities in the CSV file format. Log file is located at $path/$logfileName

#It verifies if the script is already running. If it is, then it doesn't run again.

#Code insipired from https://social.technet.microsoft.com/Forums/en-US/4d257c80-557a-4625-aad3-f2aac6e9a1bd/get-active-window-info?forum=winserverpowershell
#

#How to Run:

#Use the following command line

#powershell -windowstyle hidden -file AppActivity.ps1

#
