#!/bin/sh
#
# Microsoft AutoUpdate Trigger for Jamf Pro
# Script Version 1.7
#
## Copyright (c) 2020 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary 
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com 

# IT Admin constants for application path
PATH_WORD="/Applications/Microsoft Word.app"
PATH_EXCEL="/Applications/Microsoft Excel.app"
PATH_POWERPOINT="/Applications/Microsoft PowerPoint.app"
PATH_OUTLOOK="/Applications/Microsoft Outlook.app"
PATH_ONENOTE="/Applications/Microsoft OneNote.app"
PATH_SKYPEBUSINESS="/Applications/Skype for Business.app"
PATH_REMOTEDESKTOP="/Applications/Microsoft Remote Desktop.app"
PATH_COMPANYPORTAL="/Applications/Company Portal.app"
PATH_DEFENDER="/Applications/Microsoft Defender ATP.app"
PATH_EDGE="/Applications/Microsoft Edge.app"
PATH_TEAMS="/Applications/Microsoft Teams.app"
PATH_ONEDRIVE="/Applications/OneDrive.app"

APPID_WORD="MSWD2019"
APPID_EXCEL="XCEL2019"
APPID_POWERPOINT="PPT32019"
APPID_OUTLOOK="OPIM2019"
APPID_ONENOTE="ONMC2019"
APPID_SKYPEBUSINESS="MSFB16"
APPID_REMOTEDESKTOP="MSRD10"
APPID_COMPANYPORTAL="IMCP01"
APPID_DEFENDER="WDAV00"
APPID_EDGE="EDGE01"
APPID_TEAMS="TEAM01"
APPID_ONEDRIVE="ONDR18"

# Function to check whether MAU 3.18 or later command-line updates are available
function CheckMAUInstall() {
	if [ ! -e "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/msupdate" ]; then
    	echo "ERROR: MAU 3.18 or later is required!"
    	exit 1
	fi
}

# Function to check whether we are allowed to send Apple Events to MAU
function CheckAppleEvents() {
	MAURESULT=$(${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --config | /usr/bin/grep 'No result returned from Update Assistant')
	if [[ "$MAURESULT" = *"No result returned from Update Assistant"* ]]; then
    	echo "ERROR: Cannot send Apple Events to MAU. Check privacy settings"
    	exit 1
	fi
}

# Function to check whether MAU is up-to-date
function CheckMAUUpdate() {
	MAUUPDATE=$(${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --list | /usr/bin/grep 'MSau04')
	if [[ "$MAUUPDATE" = *"MSau04"* ]]; then
    	echo "Updating MAU to latest version... $MAUUPDATE"
    	echo "$(/bin/date)"
    	RESULT=$(${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --install --apps MSau04)
    	sleep 120
	fi
}

# Function to check whether its safe to close Excel because it has no open unsaved documents
function OppCloseExcel() {
	APPSTATE=$(${CMD_PREFIX}/usr/bin/pgrep "Microsoft Excel")
	if [ ! "$APPSTATE" == "" ]; then
		DIRTYDOCS=$(${CMD_PREFIX}/usr/bin/defaults read com.microsoft.Excel NumTotalBookDirty)
		if [ "$DIRTYDOCS" == "0" ]; then
			echo "$(/bin/date)"
			echo "Closing Excel as no unsaved documents are open"
			$(${CMD_PREFIX}/usr/bin/pkill -HUP "Microsoft Excel")
		fi
	fi
}

# Function to determine the logged-in state of the Mac
function DetermineLoginState() {
	# The following line is is taken from: https://erikberglund.github.io/2018/Get-the-currently-logged-in-user,-in-Bash/
	CONSOLE="$(/usr/sbin/scutil <<< "show State:/Users/ConsoleUser" | /usr/bin/awk '/Name :/ && ! /loginwindow/ { print $3 }')"
	if [ "$CONSOLE" == "" ]; then
		echo "No user currently logged in to console - using fall-back account"
        CONSOLE=$(/usr/bin/last -1 -t ttys000 | /usr/bin/awk '{print $1}')
        echo "Using account $CONSOLE for update"
		userID=$(/usr/bin/id -u "$CONSOLE")
		CMD_PREFIX="/bin/launchctl asuser $userID "
	else
    	echo "User $CONSOLE is logged in"
    		userID=$(/usr/bin/id -u "$CONSOLE")
		CMD_PREFIX="/bin/launchctl asuser $userID "
	fi
}

# Function to register an application with MAU
function RegisterApp() {
   	$(${CMD_PREFIX}/usr/bin/defaults write com.microsoft.autoupdate2 Applications -dict-add "$1" "{ 'Application ID' = '$2'; LCID = 1033 ; }")
}

# Function to flush any existing MAU sessions
function FlushDaemon() {
	$(${CMD_PREFIX}/usr/bin/defaults write com.microsoft.autoupdate.fba ForceDisableMerp -bool TRUE)
	$(${CMD_PREFIX}/usr/bin/pkill -HUP "Microsoft Update Assistant")
}

# Function to call 'msupdate' and update the target applications
function PerformUpdate() {
	echo "$(/bin/date)"
	${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --install --apps $1 --wait 600 2>/dev/null
}

## MAIN
echo "Started - $(/bin/date)"
DetermineLoginState
CheckMAUInstall
FlushDaemon
CheckAppleEvents
CheckMAUUpdate
FlushDaemon
RegisterApp "$PATH_WORD" "$APPID_WORD"
RegisterApp "$PATH_EXCEL" "$APPID_EXCEL"
RegisterApp "$PATH_POWERPOINT" "$APPID_POWERPOINT"
RegisterApp "$PATH_OUTLOOK" "$APPID_OUTLOOK"
RegisterApp "$PATH_ONENOTE" "$APPID_ONENOTE"
RegisterApp "$PATH_SKYPEBUSINESS" "$APPID_SKYPEBUSINESS"
RegisterApp "$PATH_REMOTEDESKTOP" "$APPID_REMOTEDESKTOP"
RegisterApp "$PATH_COMPANYPORTAL" "$APPID_COMPANYPORTAL"
RegisterApp "$PATH_DEFENDER" "$APPID_DEFENDER"
RegisterApp "$PATH_EDGE $APPID_EDGE"
RegisterApp "$PATH_TEAMS $APPID_TEAMS"
RegisterApp "$PATH_ONEDRIVE $APPID_ONEDRIVE"
OppCloseExcel

PerformUpdate "$APPID_WORD $APPID_EXCEL $APPID_POWERPOINT $APPID_OUTLOOK $APPID_ONENOTE $APPID_SKYPEBUSINESS $APPID_REMOTEDESKTOP $APPID_COMPANYPORTAL $APPID_DEFENDER $APPID_EDGE $APPID_TEAMS $APPID_ONEDRIVE"

echo "Finished - $(/bin/date)"

exit 0
