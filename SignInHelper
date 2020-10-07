#!/bin/sh
#set -x

TOOL_NAME="Microsoft Office for Mac Sign In Helper"
TOOL_VERSION="1.4.0"

## Copyright (c) 2020 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary 
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com

## This script is Jamf Pro compatible and can be pasted directly, without modification, into a new script window in the Jamf admin console.
## When running under Jamf Pro, no additional parameters need to be specified.

# Shows tool usage and parameters
ShowUsage() {
	echo "$TOOL_NAME - $TOOL_VERSION"
	echo "Purpose: Detects UPN of logged-on user and pre-fills Office and Skype for Business Sign In page"
	echo "Usage: $0 [--Verbose]"
	echo
	exit 0
}

# Checks to see if the script is running as root
RunningAsRoot() {
	# shellcheck disable=SC2039
	if [ "$EUID" = "0" ]; then
		echo "1"
	else
		echo "0"
	fi
}

# Returns the name of the logged-in user, which is useful if the script is running in the root context
GetLoggedInUser() {
	# The following line is courtesy of @scriptingosx - https://scriptingosx.com/2019/09/get-current-user-in-shell-scripts-on-macos/
	LOGGEDIN=$(/bin/echo "show State:/Users/ConsoleUser" | /usr/sbin/scutil | /usr/bin/awk '/Name :/&&!/loginwindow/{print $3}')
	if [ "$LOGGEDIN" = "" ]; then
		echo "0"
	else
		echo "$LOGGEDIN"
	fi
}

# HOME folder detection
GetHomeFolder() {
	HOME=$(dscl . read /Users/"$1" NFSHomeDirectory | cut -d ':' -f2 | cut -d ' ' -f2)
	if [ "$HOME" = "" ]; then
		if [ -d "/Users/$1" ]; then
			HOME="/Users/$1"
		else
			HOME=$(eval echo "~$1")
		fi
	fi
}

# Detects whether a given preference is managed
IsPrefManaged() {
	PREFKEY="$1"
	PREFDOMAIN="$2"
	MANAGED=$(/usr/bin/python -c "from Foundation import CFPreferencesAppValueIsForced; print CFPreferencesAppValueIsForced('${PREFKEY}', '${PREFDOMAIN}')")
	if [ "$MANAGED" = "True" ]; then
		echo "1"
	else
		echo "0"
	fi
}

# Detect Kerberos cache
DetectKerbCache() {
	KERB=$(${CMD_PREFIX} /usr/bin/klist 2> /dev/null)
	if [ "$KERB" = "" ]; then
		echo "0"
	else
		echo "1"
	fi
}

# Get the Kerberos principal from the cache
GetPrincipal() {
	PRINCIPAL=$(${CMD_PREFIX} /usr/bin/klist | grep -o 'Principal: .*' | cut -d : -f2 | cut -d' ' -f2 2> /dev/null)
	if [ "$PRINCIPAL" = "" ]; then
		echo "0"
	else
		echo "$PRINCIPAL"
	fi
}

# Extract account name from principal
GetAccountName() {
	PRINCIPAL="$1"
	echo "$PRINCIPAL" | cut -d @ -f1
}

# Extract domain name from principal
GetDomainName() {
	PRINCIPAL="$1"
	echo "$PRINCIPAL" | cut -d @ -f2
}

# Get the defaultNamingContext from LDAP
GetDefaultNamingContext() {
	DOMAIN="$1"
	DOMAINNC=$(${CMD_PREFIX} /usr/bin/ldapsearch -H "ldap://$DOMAIN" -LLL -b '' -s base defaultNamingContext | grep -o 'defaultNamingContext:.*' | cut -d : -f2 | cut -d' ' -f2 2> /dev/null)
	if [ "$DOMAINNC" = "" ]; then
		echo "0"
	else
		echo "$DOMAINNC"
	fi
}

# Get the UPN for the user account
GetUPN() {
	DOMAIN="$1"
	NAMESPACE="$2"
	ACCOUNT="$3"
	UPN=$(${CMD_PREFIX} /usr/bin/ldapsearch -H "ldap://$DOMAIN" -LLL -b "$NAMESPACE" -s sub samAccountName="$ACCOUNT" userPrincipalName | grep -o 'userPrincipalName:.*' | cut -d : -f2 | cut -d' ' -f2 2> /dev/null)
	if [ "$UPN" = "" ]; then
		echo "0"
	else
		echo "$UPN"
	fi
}

# Get the displayName for the user account
GetDisplayName() {
	DOMAIN="$1"
	NAMESPACE="$2"
	ACCOUNT="$3"
	DISPLAYNAME=$(${CMD_PREFIX} /usr/bin/ldapsearch -H "ldap://$DOMAIN" -LLL -b "$NAMESPACE" -s sub samAccountName="$ACCOUNT" displayName | grep -o 'displayName:.*' | cut -d : -f2 | cut -c 2- 2> /dev/null)
	if [ "$DISPLAYNAME" = "" ]; then
		echo "0"
	else
		echo "$DISPLAYNAME"
	fi
}

# Parse initials from the displayName
GetInitials() {
	DISPLAYNAME="$1"
	FIRST=$(echo "${DISPLAYNAME}" | cut -d ' ' -f1 | cut -c 1)
	SECOND=$(echo "${DISPLAYNAME}" | cut -d ' ' -f2 | cut -c 1)
	THIRD=$(echo "${DISPLAYNAME}" | cut -d ' ' -f3 | cut -c 1)
	INITIALS="${FIRST}${SECOND}${THIRD}"
	if [ "$INITIALS" = "" ]; then
		echo "0"
	else
		echo "$INITIALS"
	fi
}

# Set Sign In keys
SetPrefill() {
	UPN="$1"
	SetPrefillOffice "$UPN"
	SetPrefillSkypeForBusiness "$UPN"
	### Comment out the next line if you don't want to enable automatic sign in, or you are setting it separately in a Configuration Profile
	SetAutoSignIn
}

# Set Home Realm Discovery for Office apps
SetPrefillOffice() {
	UPN="$1"
	KEYMANAGED=$(IsPrefManaged "OfficeActivationEmailAddress" "com.microsoft.office")
	if [ "$KEYMANAGED" = "1" ]; then
		echo ">>ERROR - Cannot override managed preference 'OfficeActivationEmailAddress'"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write com.microsoft.office OfficeActivationEmailAddress -string "${UPN}"; then
			echo ">>SUCCESS - Set 'OfficeActivationEmailAddress' to ${UPN}"
		else
			echo ">>ERROR - Did not set value for 'OfficeActivationEmailAddress'"
			exit 1
		fi
	fi
}

# Set Office Automatic Office Sign In
SetAutoSignIn() {
	KEYMANAGED=$(IsPrefManaged "OfficeAutoSignIn" "com.microsoft.office")
	if [ "$KEYMANAGED" = "1" ]; then
		echo ">>WARNING - Cannot override managed preference 'OfficeAutoSignIn'"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write com.microsoft.office OfficeAutoSignIn -bool true; then
			echo ">>SUCCESS - Set 'OfficeAutoSignIn' to TRUE"
		else
			echo ">>ERROR - Did not set value for 'OfficeAutoSignIn'"
			exit 1
		fi
	fi
}

# Set Office Username and Initials
SetOfficeUser() {
	USERNAME="$1"
	INITIALS="$2"
	if [ "$USERNAME" = "" ]; then
		echo ">>WARNING - Cannot set Office user name"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write "${HOME}/Library/Group Containers/UBF8T346G9.Office/MeContact.plist" Name -string "${USERNAME}"; then
			echo ">>SUCCESS - Set Office user name to '${USERNAME}'"
		else
			echo ">>WARNING - Did not set Office user name'"
		fi
	fi
	if [ "$INITIALS" = "" ]; then
		echo ">>WARNING - Cannot set Office user initials"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write "${HOME}/Library/Group Containers/UBF8T346G9.Office/MeContact.plist" Initials -string "${INITIALS}"; then
			echo ">>SUCCESS - Set Office user initials to '${INITIALS}'"
		else
			echo ">>WARNING - Did not set Office user initials'"
		fi
	fi
}

# Set Skype for Business Sign In
SetPrefillSkypeForBusiness() {
	UPN="$1"
	KEYMANAGED=$(IsPrefManaged "userName" "com.microsoft.SkypeForBusiness")
	if [ "$KEYMANAGED" = "1" ]; then
		echo ">>ERROR - Cannot override managed preference 'userName'"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write com.microsoft.SkypeForBusiness userName -string "${UPN}"; then
			echo ">>SUCCESS - Set 'userName' to ${UPN}"
		else
			echo ">>ERROR - Did not set value for 'userName'"
			exit 1
		fi
	fi
	SIP="$1"
	KEYMANAGED=$(IsPrefManaged "sipAddress" "com.microsoft.SkypeForBusiness")
	if [ "$KEYMANAGED" = "1" ]; then
		echo ">>ERROR - Cannot override managed preference 'sipAddress'"
	else
		if ${CMD_PREFIX} /usr/bin/defaults write com.microsoft.SkypeForBusiness sipAddress -string "${SIP}"; then
			echo ">>SUCCESS - Set 'sipAddress' to ${SIP}"
		else
			echo ">>ERROR - Did not set value for 'sipAddress'"
			exit 1
		fi
	fi
}

# Detect Domain Join
DetectDomainJoin() {
	DSCONFIGAD=$(${CMD_PREFIX} /usr/sbin/dsconfigad -show)
	if [ "$DSCONFIGAD" = "" ]; then
		echo "0"
	else
		echo "1"
	fi
}

# Detect Jamf presence
DetectJamf() {
	if [ -e "/Library/Preferences/com.jamfsoftware.jamf.plist" ]; then
		echo "1"
	else
		echo "0"
	fi
}

# Detect Jamf Connect presence
DetectJamfConnect() {
	JAMFC=$(${CMD_PREFIX} /usr/bin/defaults read com.jamf.connect.state 2> /dev/null)
	if [ "$JAMFC" = "" ]; then
		echo "0"
	else
		echo "1"
	fi
}
# Get the DisplayName from Jamf Connects preference cache
GetDisplayNamefromJamfConnect() {
	JCDISPLAYNAME=$(${CMD_PREFIX} /usr/bin/defaults read com.jamf.connect.state DisplayName 2> /dev/null)
	if [ "$JCDISPLAYNAME" = "" ]; then
		echo "0"
	else
		echo "$JCDISPLAYNAME"
	fi
}

# Detect NoMAD presence
DetectNoMAD() {
	NOMAD=$(${CMD_PREFIX} /usr/bin/defaults read com.trusourcelabs.NoMAD 2> /dev/null)
	if [ "$NOMAD" = "" ]; then
		echo "0"
	else
		echo "1"
	fi
}

# Get the UPN from NoMAD's preference cache
GetUPNfromNoMAD() {
	NMUPN=$(${CMD_PREFIX} /usr/bin/defaults read com.trusourcelabs.NoMAD UserUPN 2> /dev/null)
	if [ "$NMUPN" = "" ]; then
		echo "0"
	else
		echo "$NMUPN"
	fi
}

# Get the DisplayName from NoMAD's preference cache
GetDisplayNamefromNoMAD() {
	NMDISPLAYNAME=$(${CMD_PREFIX} /usr/bin/defaults read com.trusourcelabs.NoMAD DisplayName 2> /dev/null)
	if [ "$NMDISPLAYNAME" = "" ]; then
		echo "0"
	else
		echo "$NMDISPLAYNAME"
	fi
}

# Detect Enterprise Connect presence
DetectEnterpriseConnect() {
	EC=$(${CMD_PREFIX} /usr/bin/defaults read com.apple.Enterprise-Connect 2> /dev/null)
	if [ "$EC" = "" ]; then
		echo "0"
	else
		echo "1"
	fi
}

# Get the UPN from Enterprise Connect (code courtesy of Dennis Browning)
GetUPNfromEnterpriseConnect() {
	ECUPN=$(${CMD_PREFIX} "/Applications/Enterprise Connect.app/Contents/SharedSupport/eccl" -a userPrincipalName | awk '/userPrincipalName:/{print $NF}' 2> /dev/null)
	if [ "$ECUPN" = "" ]; then
		echo "0"
	else
		echo "$ECUPN"
	fi
}

# Evaluate command-line arguments
while [ $# -gt 0 ]; do
	key="$1"
	case "$key" in
		--Help|-h|--help)
		ShowUsage
		exit 0
		shift # past argument
		;;
		--Verbose|-v|--verbose)
		set -x
		shift # past argument
		;;
	esac
	shift # past argument or value
done

## Main
# Determine whether we need to use a sudo -u prefix when running commands
# NOTE: CMD_PREFIX is intentionally implemented as a global variable
CMD_PREFIX=""
ROOTLOGON=$(RunningAsRoot)
if [ "$ROOTLOGON" = "1" ]; then
	CURRENTUSER=$(GetLoggedInUser)
	GetHomeFolder "$CURRENTUSER"
	if [ ! "$CURRENTUSER" = "0" ]; then
		echo ">>INFO - Script is running in the root security context - running commands as user: $CURRENTUSER"
		CMD_PREFIX="/usr/bin/sudo -u ${CURRENTUSER}"
	else
		echo ">>ERROR - Could not obtain the logged in user name"
		exit 1
	fi
fi

# Detect Active Directory connection style
DJ=$(DetectDomainJoin)
if [ "$DJ" = "1" ]; then
	echo ">>INFO - Detected that this machine is domain joined"
fi
NM=$(DetectNoMAD)
if [ "$NM" = "1" ]; then
	echo ">>INFO - Detected that this machine is running NoMAD"
fi
EC=$(DetectEnterpriseConnect)
if [ "$EC" = "1" ]; then
	echo ">>INFO - Detected that this machine is running Enterprise Connect"
fi
JC=$(DetectJamfConnect)
if [ "$JC" = "1" ]; then
	echo ">>INFO - Detected that this machine is running Jamf Connect"
fi

# Find out if a Kerberos principal and ticket is present
UPN="0"
KERBCACHE=$(DetectKerbCache)
if [ "$KERBCACHE" = "1" ]; then
	echo ">>INFO - Detected Kerberos cache"
	PRINCIPAL=$(GetPrincipal)
	if [ ! "$PRINCIPAL" = "0" ]; then
		echo ">>INFO - Detected Kerberos principal: $PRINCIPAL"
		# Get the account and domain name
		ACCOUNT=$(GetAccountName "$PRINCIPAL")
		DOMAIN=$(GetDomainName "$PRINCIPAL")
		# Find the default naming context for Active Directory
		NAMESPACE=$(GetDefaultNamingContext "$DOMAIN")
		if [ ! "$NAMESPACE" = "0" ]; then
			echo ">>INFO - Detected naming context: $NAMESPACE"
			# Now to get the UPN
			UPN=$(GetUPN "$DOMAIN" "$NAMESPACE" "$ACCOUNT")
			if [ ! "$UPN" = "0" ]; then
				echo ">>INFO - Found UPN: $UPN"
				SetPrefill "$UPN"
				DISPLAYNAME=$(GetDisplayName "$DOMAIN" "$NAMESPACE" "$ACCOUNT")
				if [ ! "$DISPLAYNAME" = "0" ]; then
					echo ">>INFO - Found DisplayName: $DISPLAYNAME"
					INITIALS=$(GetInitials "$DISPLAYNAME")
					SetOfficeUser "$DISPLAYNAME" "$INITIALS"
				fi
				exit 0
			else
				echo ">>WARNING - Could not find UPN"
			fi
		else
			echo ">>WARNING - Could not retrieve naming context"
		fi
	else
		echo ">>WARNING - Could not retrieve principal"
	fi
else
	echo ">>WARNING - No Kerberos cache present"
fi

# If we haven't got a UPN yet, see if we can get it from NoMAD's cache
if [ "$UPN" = "0" ] && [ "$NM" = "1" ]; then
	UPN=$(GetUPNfromNoMAD)
	if [ ! "$UPN" = "0" ]; then
		echo ">>INFO - Found UPN from NoMAD: $UPN"
		SetPrefill "$UPN"
		DISPLAYNAME=$(GetDisplayNamefromNoMAD)
			if [ ! "$DISPLAYNAME" = "0" ]; then
				echo ">>INFO - Found DisplayName from NoMAD: $DISPLAYNAME"
				INITIALS=$(GetInitials "$DISPLAYNAME")
				SetOfficeUser "$DISPLAYNAME" "$INITIALS"
			fi
		exit 0
	else
		echo ">>WARNING - Could not retrieve UPN from NoMAD"
	fi
fi

# If we still haven't got a UPN, see if we can get it from Enterprise Connect (code courtesy of Dennis Browning)
if [ "$UPN" = "0" ] && [ "$EC" = "1" ]; then
	UPN=$(GetUPNfromEnterpriseConnect)
	if [ ! "$UPN" = "0" ]; then
		echo ">>INFO - Found UPN from Enterprise Connect: $UPN"
		SetPrefill "$UPN"
		exit 0
	else
		echo ">>WARNING - Could not retrieve UPN from Enterprise Connect"
	fi
fi

# Check Jamf Connect for Display Name
if [ "$UPN" = "0" ] && [ "$JC" = "1" ]; then
	UPN=$(GetDisplayNamefromJamfConnect)
	if [ ! "$UPN" = "0" ]; then
		echo ">>INFO - Found UPN under DisplayName from Jamf Connect: $UPN"
		SetPrefill "$UPN"
		exit 0
	else
		echo ">>WARNING - Could not retrieve UPN from Jamf Connect"
	fi
fi

# If we still haven't got a UPN yet, show an error
if [ "$UPN" = "0" ]; then
	echo ">>ERROR - Could not detect UPN"
	exit 1
fi

exit 0