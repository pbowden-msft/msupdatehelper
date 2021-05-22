#!/bin/sh
#
# Microsoft AutoUpdate Helper for Jamf Pro
# Script Version 1.4
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

# IT Admin constants for which applications to update [set to true or false as required]
UPDATE_WORD="true"
UPDATE_EXCEL="true"
UPDATE_POWERPOINT="true"
UPDATE_OUTLOOK="true"
UPDATE_ONENOTE="true"
UPDATE_SKYPEBUSINESS="true"
UPDATE_REMOTEDESKTOP="true"
UPDATE_COMPANYPORTAL="true"

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

# Function to enable debug logging
function Debug() {
    if [ "$OVERRIDE_DEBUG" == "true" ] || [ "$OVERRIDE_DEBUG" == "TRUE" ] || [ "$OVERRIDE_DEBUG" == "True" ] || [ "$OVERRIDE_DEBUG" == "YES" ] || [ "$OVERRIDE_DEBUG" == "yes" ] || [ "$OVERRIDE_DEBUG" == "Yes" ]; then
        LOG=$(date; echo "$1")
        echo "$LOG"
    fi
}

# Harvest script parameter overrides
OVERRIDE_DEBUG="$4"
OVERRIDE_WORD="$5"
Debug "OVERRIDE_WORD: $5"
OVERRIDE_EXCEL="$6"
Debug "OVERRIDE_EXCEL: $6"
OVERRIDE_POWERPOINT="$7"
Debug "OVERRIDE_POWERPOINT: $7"
OVERRIDE_OUTLOOK="$8"
Debug "OVERRIDE_OUTLOOK: $8"
OVERRIDE_SKYPEBUSINESS="$9"
Debug "OVERRIDE_SKYPEBUSINESS: $9"
OVERRIDE_ONENOTE="${10}"
Debug "OVERRIDE_ONENOTE: ${10}"
OVERRIDE_REMOTEDESKTOP="${11}"
Debug "OVERRIDE_REMOTEDESKTOP: ${11}"

# Function to evaluate app update override
function GetUpdateOverride() {
    if [ ! "$1" = "" ]; then
        local UPDATE_FIELD1=$(echo "$1" | cut -d '@' -f1)
        if [ "$UPDATE_FIELD1" == "TRUE" ] || [ "$UPDATE_FIELD1" == "true" ] || [ "$UPDATE_FIELD1" == "True" ] || [ "$UPDATE_FIELD1" == "YES" ] || [ "$UPDATE_FIELD1" == "yes" ] || [ "$UPDATE_FIELD1" == "Yes" ]; then
            echo "true"
        elif [ "$UPDATE_FIELD1" == "FALSE" ] || [ "$UPDATE_FIELD1" == "false" ] || [ "$UPDATE_FIELD1" == "False" ] || [ "$UPDATE_FIELD1" == "NO" ] || [ "$UPDATE_FIELD1" == "no" ] || [ "$UPDATE_FIELD1" == "No" ]; then
            echo "false"
        fi
    else
        echo "$2"
    fi
}

# Function to evaluate app version override
function GetVersionOverride() {
    if [ ! "$1" = "" ]; then
        local UPDATE_FIELD2=$(echo "$1" | cut -d '@' -f2)
        if [ "$UPDATE_FIELD2" == "TRUE" ] || [ "$UPDATE_FIELD2" == "true" ]  || [ "$UPDATE_FIELD2" == "True" ] || [ "$UPDATE_FIELD2" == "YES" ] || [ "$UPDATE_FIELD2" == "yes" ]  || [ "$UPDATE_FIELD2" == "Yes" ] || [ "$UPDATE_FIELD2" == "FALSE" ] || [ "$UPDATE_FIELD2" == "false" ] || [ "$UPDATE_FIELD2" == "False" ]  || [ "$UPDATE_FIELD2" == "NO" ] || [ "$UPDATE_FIELD2" == "no" ] || [ "$UPDATE_FIELD2" == "No" ] ; then
            echo "$2"
        else
            echo "$UPDATE_FIELD2"
        fi
    else
        echo "$2"
    fi
}

# Function to parse script parameter overrides
function GetOverrides() {
    UPDATE_WORD=$(GetUpdateOverride "$OVERRIDE_WORD" "$UPDATE_WORD")
    Debug "Resolved UPDATE_WORD: $UPDATE_WORD"
    VERSION_WORD=$(GetVersionOverride "$OVERRIDE_WORD" "$VERSION_WORD")
    Debug "Resolved VERSION_WORD: $VERSION_WORD"

    UPDATE_EXCEL=$(GetUpdateOverride "$OVERRIDE_EXCEL" "$UPDATE_EXCEL")
    Debug "Resolved UPDATE_EXCEL: $UPDATE_EXCEL"
    VERSION_EXCEL=$(GetVersionOverride "$OVERRIDE_EXCEL" "$VERSION_EXCEL")
    Debug "Resolved VERSION_EXCEL: $VERSION_EXCEL"
    
    UPDATE_POWERPOINT=$(GetUpdateOverride "$OVERRIDE_POWERPOINT" "$UPDATE_POWERPOINT")
    Debug "Resolved UPDATE_POWERPOINT: $UPDATE_POWERPOINT"
    VERSION_POWERPOINT=$(GetVersionOverride "$OVERRIDE_POWERPOINT" "$VERSION_POWERPOINT")
    Debug "Resolved VERSION_POWERPOINT: $VERSION_POWERPOINT"

    UPDATE_OUTLOOK=$(GetUpdateOverride "$OVERRIDE_OUTLOOK" "$UPDATE_OUTLOOK")
    Debug "Resolved UPDATE_OUTLOOK: $UPDATE_OUTLOOK"
    VERSION_OUTLOOK=$(GetVersionOverride "$OVERRIDE_OUTLOOK" "$VERSION_OUTLOOK")
    Debug "Resolved VERSION_OUTLOOK: $VERSION_OUTLOOK"

    UPDATE_SKYPEBUSINESS=$(GetUpdateOverride "$OVERRIDE_SKYPEBUSINESS" "$UPDATE_SKYPEBUSINESS")
    Debug "Resolved UPDATE_SKYPEBUSINESS: $UPDATE_SKYPEBUSINESS"
    VERSION_SKYPEBUSINESS=$(GetVersionOverride "$OVERRIDE_SKYPEBUSINESS" "$VERSION_SKYPEBUSINESS")
    Debug "Resolved VERSION_SKYPEBUSINESS: $VERSION_SKYPEBUSINESS"
    
    UPDATE_ONENOTE=$(GetUpdateOverride "$OVERRIDE_ONENOTE" "$UPDATE_ONENOTE")
    Debug "Resolved UPDATE_ONENOTE: $UPDATE_ONENOTE"
    VERSION_ONENOTE=$(GetVersionOverride "$OVERRIDE_ONENOTE" "$VERSION_ONENOTE")
    Debug "Resolved VERSION_ONENOTE: $VERSION_ONENOTE"
    
    UPDATE_REMOTEDESKTOP=$(GetUpdateOverride "$OVERRIDE_REMOTEDESKTOP" "$UPDATE_REMOTEDESKTOP")
    Debug "Resolved UPDATE_REMOTEDESKTOP: $UPDATE_REMOTEDESKTOP"
    VERSION_REMOTEDESKTOP=$(GetVersionOverride "$OVERRIDE_REMOTEDESKTOP" "$VERSION_REMOTEDESKTOP")
    Debug "Resolved VERSION_REMOTEDESKTOP: $VERSION_REMOTEDESKTOP"
}

# Function to check whether MAU 3.18 or later command-line updates are available
function CheckMAUInstall() {
	if [ ! -e "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/msupdate" ]; then
    	echo "MAU 3.18 or later is required!"
    	exit 1
	fi
}

# Function to check whether Office apps are installed
function CheckAppInstall() {
	if [ ! -e "$PATH_WORD" ]; then
    	Debug "Word is not installed"
    	UPDATE_WORD="false"
	fi
	if [ ! -e "$PATH_EXCEL" ]; then
    	Debug "Excel is not installed"
    	UPDATE_EXCEL="false"
	fi
	if [ ! -e "$PATH_POWERPOINT" ]; then
    	Debug "PowerPoint is not installed"
    	UPDATE_POWERPOINT="false"
	fi
	if [ ! -e "$PATH_OUTLOOK" ]; then
    	Debug "Outlook is not installed"
    	UPDATE_OUTLOOK="false"
	fi
	if [ ! -e "$PATH_ONENOTE" ]; then
    	Debug "OneNote is not installed"
    	UPDATE_ONENOTE="false"
	fi
	if [ ! -e "$PATH_SKYPEBUSINESS" ]; then
    	Debug "Skype for Business is not installed"
    	UPDATE_SKYPEBUSINESS="false"
	fi
	if [ ! -e "$PATH_REMOTEDESKTOP" ]; then
    	Debug "Remote Desktop is not installed"
    	UPDATE_REMOTEDESKTOP="false"
	fi
	if [ ! -e "$PATH_COMPANYPORTAL" ]; then
    	Debug "Company Portal is not installed"
    	UPDATE_COMPANYPORTAL="false"
	fi
}

# Function to determine the logged-in state of the Mac
function DetermineLoginState() {
	# The following line is courtesy of @macmule - https://macmule.com/2014/11/19/how-to-get-the-currently-logged-in-user-in-a-more-apple-approved-way/
	CONSOLE=$(/usr/bin/python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");')
	if [ "$CONSOLE" == "" ]; then
    	echo "No user logged in"
		CMD_PREFIX=""
	else
    	echo "User $CONSOLE is logged in"
    	userID=$(/usr/bin/id -u "$CONSOLE")
	CMD_PREFIX="/bin/launchctl asuser $userID "
	fi
	Debug "Resolved CMD_PREFIX: $CMD_PREFIX"
}

# Function to set target version for app
function SetTargetVersion() {
	if [ "$1" == "LATEST" ] || [ "$1" == "latest" ] || [ "$1" == "" ]; then
		TARGET_VERSION=""
	else
		TARGET_VERSION="--version ${1}"
	fi
	Debug "Final TARGET_VERSION: $TARGET_VERSION"
}

# Function to register an application with MAU
function RegisterApp() {
   	Debug "RegisterApp: Params - $1 $2"
   	$(${CMD_PREFIX}defaults write com.microsoft.autoupdate2 Applications -dict-add "$1" "{ 'Application ID' = '$2'; LCID = 1033 ; }")
}

# Function to call 'msupdate' and update the target application
function PerformUpdate() {
    Debug "PerformUpdate: ${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --install --apps $1 $2 --wait 600 2>/dev/null"
	${CMD_PREFIX}/Library/Application\ Support/Microsoft/MAU2.0/Microsoft\ AutoUpdate.app/Contents/MacOS/msupdate --install --apps $1 $2 --wait 600 2>/dev/null
}

## MAIN
CheckMAUInstall
GetOverrides
CheckAppInstall
DetermineLoginState

if [ "$UPDATE_WORD" == "true" ]; then
	Debug "Going for Word update"
	RegisterApp "$PATH_WORD" "MSWD15"
	SetTargetVersion "$VERSION_WORD"
	PerformUpdate "MSWD15" "$TARGET_VERSION"
else
	Debug "Update for Word disabled"
fi
if [ "$UPDATE_EXCEL" == "true" ]; then
	Debug "Going for Excel update"
	RegisterApp "$PATH_EXCEL" "XCEL15"
	SetTargetVersion "$VERSION_EXCEL"
	PerformUpdate "XCEL15" "$TARGET_VERSION"
else
	Debug "Update for Excel disabled"
fi
if [ "$UPDATE_POWERPOINT" == "true" ]; then
	Debug "Going for PowerPoint update"
	RegisterApp "$PATH_POWERPOINT" "PPT315"
	SetTargetVersion "$VERSION_POWERPOINT"
	PerformUpdate "PPT315" "$TARGET_VERSION"
else
	Debug "Update for PowerPoint disabled"
fi
if [ "$UPDATE_OUTLOOK" == "true" ]; then
	Debug "Going for Outlook update"
	RegisterApp "$PATH_OUTLOOK" "OPIM15"
	SetTargetVersion "$VERSION_OUTLOOK"
	PerformUpdate "OPIM15" "$TARGET_VERSION"
else
	Debug "Update for Outlook disabled"
fi
if [ "$UPDATE_ONENOTE" == "true" ]; then
	Debug "Going for OneNote update"
	RegisterApp "$PATH_ONENOTE" "ONMC15"
	SetTargetVersion "$VERSION_ONENOTE"
	PerformUpdate "ONMC15" "$TARGET_VERSION"
else
	Debug "Update for OneNote disabled"
fi
if [ "$UPDATE_SKYPEBUSINESS" == "true" ]; then
	Debug "Going for SfB update"
	RegisterApp "$PATH_SKYPEBUSINESS" "MSFB16"
	SetTargetVersion "$VERSION_SKYPEBUSINESS"
	PerformUpdate "MSFB16" "$TARGET_VERSION"
else
	Debug "Update for SfB disabled"
fi
if [ "$UPDATE_REMOTEDESKTOP" == "true" ]; then
	Debug "Going for Remote Desktop update"
	RegisterApp "$PATH_REMOTEDESKTOP" "MSRD10"
	SetTargetVersion "$VERSION_REMOTEDESKTOP"
	PerformUpdate "MSRD10" "$TARGET_VERSION"
else
	Debug "Update for Remote Desktop disabled"
fi
if [ "$UPDATE_COMPANYPORTAL" == "true" ]; then
	Debug "Going for Company Portal update"
	RegisterApp "$PATH_COMPANYPORTAL" "IMCP01"
	SetTargetVersion "$VERSION_COMPANYPORTAL"
	PerformUpdate "IMCP01" "$TARGET_VERSION"
else
	Debug "Update for Company Portal disabled"
fi

exit 0
