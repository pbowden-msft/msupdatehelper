#!/bin/sh
#
# Microsoft AutoUpdate 4.0 Helper for Jamf Pro
# Script Version 1.0
#
## Copyright (c) 2018 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary 
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com

# IT Admin constants for application control [true = IT controls updates, false = Users control updates]
MANAGED_UPDATES=false
RUN_RECON=false

# IT Admin constants for which applications to update [set to true or false as required]
UPDATE_WORD=true
UPDATE_EXCEL=true
UPDATE_POWERPOINT=true
UPDATE_OUTLOOK=true
UPDATE_ONENOTE=true
UPDATE_SKYPEBUSINESS=true
UPDATE_REMOTEDESKTOP=true
UPDATE_COMPANYPORTAL=true

# IT Admin constants for application target version [set to "latest" to get latest update, or specific build number, such as "15.41.17120500"]
VERSION_WORD="latest"
VERSION_EXCEL="latest"
VERSION_POWERPOINT="latest"
VERSION_OUTLOOK="latest"
VERSION_ONENOTE="latest"
VERSION_SKYPEBUSINESS="latest"
VERSION_REMOTEDESKTOP="latest"
VERSION_COMPANYPORTAL="latest"

# IT Admin constants for application path
PATH_WORD="/Applications/Microsoft Word.app"
PATH_EXCEL="/Applications/Microsoft Excel.app"
PATH_POWERPOINT="/Applications/Microsoft PowerPoint.app"
PATH_OUTLOOK="/Applications/Microsoft Outlook.app"
PATH_ONENOTE="/Applications/Microsoft OneNote.app"
PATH_SKYPEBUSINESS="/Applications/Skype for Business.app"
PATH_REMOTEDESKTOP="/Applications/Microsoft Remote Desktop.app"
PATH_COMPANYPORTAL="/Applications/Company Portal.app"

# Harvest script parameter overrides
OVERRIDE_MANAGED="$4"
OVERRIDE_RECON="$5"
OVERRIDE_WORD="$6"
OVERRIDE_EXCEL="$7"
OVERRIDE_POWERPOINT="$8"
OVERRIDE_OUTLOOK="$9"
OVERRIDE_ONENOTE="$10"
OVERRIDE_SKYPEBUSINESS="$11"

# Function to parse script parameter overrides
function GetOverrides() {
    if [ "$OVERRIDE_MANAGED" = "TRUE" ] || [ "$OVERRIDE_MANAGED" = "true" ] || [ "$OVERRIDE_MANAGED" = "YES" ] || [ "$OVERRIDE_MANAGED" = "yes" ]; then
        MANAGED_UPDATES=true
    elif [ "$OVERRIDE_MANAGED" = "FALSE" ] || [ "$OVERRIDE_MANAGED" = "false" ] || [ "$OVERRIDE_MANAGED" = "NO" ] || [ "$OVERRIDE_MANAGED" = "no" ]; then
        MANAGED_UPDATES=false
    fi
    if [ "$OVERRIDE_RECON" = "TRUE" ] || [ "$OVERRIDE_RECON" = "true" ] || [ "$OVERRIDE_RECON" = "YES" ] || [ "$OVERRIDE_RECON" = "yes" ]; then
        RUN_RECON=true
    elif [ "$OVERRIDE_RECON" = "FALSE" ] || [ "$OVERRIDE_RECON" = "false" ] || [ "$OVERRIDE_RECON" = "NO" ] || [ "$OVERRIDE_RECON" = "no" ]; then
        RUN_RECON=false
    fi
    if [ ! "$OVERRIDE_WORD" = "" ]; then
        UPDATE_WORD=$(echo "$OVERRIDE_WORD" | cut -d '@' -f1)
        VERSION_WORD=$(echo "$OVERRIDE_WORD" | cut -d '@' -f2)
    fi
    if [ ! "$OVERRIDE_EXCEL" = "" ]; then
        UPDATE_EXCEL=$(echo "$OVERRIDE_EXCEL" | cut -d '@' -f1)
        VERSION_EXCEL=$(echo "$OVERRIDE_EXCEL" | cut -d '@' -f2)
    fi
    if [ ! "$OVERRIDE_POWERPOINT" = "" ]; then
        UPDATE_POWERPOINT=$(echo "$OVERRIDE_POWERPOINT" | cut -d '@' -f1)
        VERSION_POWERPOINT=$(echo "$OVERRIDE_POWERPOINT" | cut -d '@' -f2)
    fi
    if [ ! "$OVERRIDE_OUTLOOK" = "" ]; then
        UPDATE_OUTLOOK=$(echo "$OVERRIDE_OUTLOOK" | cut -d '@' -f1)
        VERSION_OUTLOOK=$(echo "$OVERRIDE_OUTLOOK" | cut -d '@' -f2)
    fi
    if [ ! "$OVERRIDE_ONENOTE" = "" ]; then
        UPDATE_ONENOTE=$(echo "$OVERRIDE_ONENOTE" | cut -d '@' -f1)
        VERSION_ONENOTE=$(echo "$OVERRIDE_ONENOTE" | cut -d '@' -f2)
    fi
    if [ ! "$OVERRIDE_SKYPEBUSINESS" = "" ]; then
        UPDATE_SKYPEBUSINESS=$(echo "$OVERRIDE_SKYPEBUSINESS" | cut -d '@' -f1)
        VERSION_SKYPEBUSINESS=$(echo "$OVERRIDE_SKYPEBUSINESS" | cut -d '@' -f2)
    fi
}

# Function to check whether MAU 4.0 command-line updates are available
function CheckMAUInstall() {
	if [ ! -e "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/msupdate" ]; then
    	echo "MAU 4.0 is not installed"
    	exit 1
	fi
}

# Function to determine the logged-in state of the Mac
function DetermineLoginState() {
	CONSOLE=$(stat -f%Su /dev/console)
	if [ "$CONSOLE" == "root" ]; then
    	echo "No user logged in"
		CMD_PREFIX=""
	else
    	echo "User $CONSOLE is logged in"
    	CMD_PREFIX="sudo -u $CONSOLE "
	fi
}

# Function to set target version for app
function SetTargetVersion() {
	if [ "$1" == "LATEST" ] || [ "$1" == "latest" ] || [ "$1" == "" ]; then
		TARGET_VERSION=""
	else
		TARGET_VERSION="--version ${1}"
	fi
}

# Function to set managed update preferences
function SetManagedPrefs() {
	if [ $MANAGED_UPDATES = true ]; then
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 HowToCheck -string 'Manual')
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 DisableInsiderCheckbox -bool TRUE)
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 EnableCheckForUpdatesButton -bool FALSE)
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 StartDaemonOnAppLaunch -bool FALSE)
    else
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 HowToCheck -string 'AutomaticDownload')
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 DisableInsiderCheckbox -bool FALSE)
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 EnableCheckForUpdatesButton -bool TRUE)
	    $(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 StartDaemonOnAppLaunch -bool TRUE)    
    fi
}

# Function to register an application with MAU
function RegisterApp() {
	if [ -d "$1" ]; then
    	$(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 Applications -dict-add "$1" "{ 'Application ID' = '$2'; LCID = 1033 ; }")
    fi
}

# Function to call 'msupdate' and update the target application
function PerformUpdate() {
	${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --install --apps $1 $2 --wait 300 2>/dev/null
}

# Function to run Jamf's Recon to update the inventory
function RunRecon() {
    if [ $RUN_RECON = true ]; then
    	$(sudo jamf recon)
    fi
}

## MAIN
CheckMAUInstall
GetOverrides
DetermineLoginState
SetManagedPrefs

if [ $UPDATE_WORD = true ]; then
	RegisterApp "$PATH_WORD" "MSWD15"
	SetTargetVersion "$VERSION_WORD"
	PerformUpdate "MSWD15" "$TARGET_VERSION"
fi
if [ $UPDATE_EXCEL = true ]; then
	RegisterApp "$PATH_EXCEL" "XCEL15"
	SetTargetVersion "$VERSION_EXCEL"
	PerformUpdate "XCEL15" "$TARGET_VERSION"
fi
if [ $UPDATE_POWERPOINT = true ]; then
	RegisterApp "$PATH_POWERPOINT" "PPT315"
	SetTargetVersion "$VERSION_POWERPOINT"
	PerformUpdate "PPT315" "$TARGET_VERSION"
fi
if [ $UPDATE_OUTLOOK = true ]; then
	RegisterApp "$PATH_OUTLOOK" "OPIM15"
	SetTargetVersion "$VERSION_OUTLOOK"
	PerformUpdate "OPIM15" "$TARGET_VERSION"
fi
if [ $UPDATE_ONENOTE = true ]; then
	RegisterApp "$PATH_ONENOTE" "ONMC15"
	SetTargetVersion "$VERSION_ONENOTE"
	PerformUpdate "ONMC15" "$TARGET_VERSION"
fi
if [ $UPDATE_SKYPEBUSINESS = true ]; then
	RegisterApp "$PATH_SKYPEBUSINESS" "MSFB16"
	SetTargetVersion "$VERSION_SKYPEBUSINESS"
	PerformUpdate "MSFB16" "$TARGET_VERSION"
fi
if [ $UPDATE_REMOTEDESKTOP = true ]; then
	RegisterApp "$PATH_REMOTEDESKTOP" "MSRD10"
	SetTargetVersion "$VERSION_REMOTEDESKTOP"
	PerformUpdate "MSRD10" "$TARGET_VERSION"
fi
if [ $UPDATE_COMPANYPORTAL = true ]; then
	RegisterApp "$PATH_COMPANYPORTAL" "IMCP01"
	SetTargetVersion "$VERSION_COMPANYPORTAL"
	PerformUpdate "IMCP01" "$TARGET_VERSION"
fi

RunRecon

exit 0
