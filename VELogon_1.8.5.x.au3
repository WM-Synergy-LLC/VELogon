#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SynergyNew.ico
#AutoIt3Wrapper_Outfile=VELogon.exe
#AutoIt3Wrapper_Res_Description=Configure Infor client user settings
#AutoIt3Wrapper_Res_Fileversion=1.8.5.88
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2008-2014 Synergy Resources, LLC
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=Copyright|Copyright © 2008-2014 Synergy Resources, LLC, Central Islip, NY. All rights reserved.
#AutoIt3Wrapper_Res_Field=Company|Synergy Resources, LLC, Central Islip, NY
#AutoIt3Wrapper_Res_Field=Author|Craig D. Gunst
#AutoIt3Wrapper_Res_Field=Email|More.Info@SynergyResources.net
#AutoIt3Wrapper_Res_Field=Original filename|VELogon.exe
#AutoIt3Wrapper_Res_Field=CompiledScript|AutoIt v3 Script
#AutoIt3Wrapper_Res_Field=CompiledDateTime|%date% - %time%
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Licence Agreement:|License Agreement: -------------------------------------------- This software is the intellectual property of Synergy Resources, LLC, Central Islip, NY and as such, is not to be distributed via any website domain or any other media without the prior written approval of Synergy Resources, LLC, Central Islip, NY. -------------------------------------------- In no event shall Synergy Resources, LLC, Central Islip, NY be liable to any party for direct, indirect, special, incidental, or consequential damages, including lost profits, arising out of the use of this software, even if Synergy Resources, LLC, Central Islip, NY has been advised of the possibility of such damage. Synergy Resources, LLC, Central Islip, NY specifically disclaims any warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The software provided hereunder is on an "as is" basis, and Synergy Resources, LLC, Central Islip, NY has no obligations to provide maintenance, support, updates, enhancements, or modifications.
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#cs -------------------------------------------------------------------------------------

	AutoIt Version: 3.2.8.1
	Author:         Craig D. Gunst (cdg)
	Craig.Gunst@SynergyResources.net

	Script Name:    VELogon.au3
	Customer:       Multiple

	Script Function:
	GPO Logon script to update client files and registry


	Copyright © 2008-2014 Synergy Resources, LLC, Central Islip, NY. All rights reserved.

	License Agreement:
	-------------------------------------------------------------------------------------
	This software is the intellectual property of Synergy Resources, LLC, Central Islip, NY
	and as such, is not to be distributed via any website domain or any other media
	without the prior written approval of Synergy Resources, LLC, Central Islip, NY.

	In no event shall Synergy Resources, LLC, Central Islip, NY be liable to any party
	for direct, indirect, special, incidental, or consequential damages, including lost
	profits, arising out of the use of this software, even if Synergy Resources, LLC,
	Central Islip, NY has been advised of the possibility of such damage. Synergy
	Resources, LLC, Central Islip, NY specifically disclaims any warranties, including,
	but not limited to, the implied warranties of merchantability and fitness for a
	particular purpose. The software provided hereunder is on an "as is" basis, and
	Synergy Resources, LLC, Central Islip, NY has no obligations to provide
	maintenance, support, updates, enhancements, or modifictions.
	-------------------------------------------------------------------------------------

	Releases:
	2008-03-19	1.2.0.1 /cdg
	- Initial Release for REMPRO
	- Added Synergy Standard AutoIt3Wrapper and Program Info section

	2008-03-20	1.2.0.2	/cdg
	- Fix minor typos in AutoIt3Wrapper

	2008-06-11	1.2.0.3 /cdg
	- Changed HKLM tell for Visual for 6.5.3
	(need a more consistant tell)

	2008-07-02	1.3.0.1 /cdg
	- Initial Release for WIRCLO
	- Changed tell for all products
	- Fixed leading/trailing space in CRM lists in .ini
	- Create blank directories in Policy Profile
	(need to do the same for Default Profile)

	2008-07-18	1.3.0.3 /cdg
	- Initial release for MEDCOR

	2008-07-18	1.3.0.4	/cdg
	- Added resolution of multiple local paths in .ini

	2008-07-25	1.3.0.5 /cdg
	- Upgrade release for RAMIND
	- Fix exit on error reading .ini

	2008-07-26	1.3.1.1	/cdg
	- Added management of the BuildInformation registry key uniqueness
	- Added build profile folder hierarchy if missing
	- Added exit if no Infor products installed (HKLM\...\Infor... exits)
	- Added recording script run info in registry
	- Added exit if cannot update ScriptFile.ini file with BuildInformation

	2008-07-27	1.3.1.2	/cdg
	- Added spec for location of VEBuildInformationLog under [VELogon]
	- Changed handling of application path to search list methood

	2008-09-06	1.3.2.1 /cdg
	- Added flags to disable various sections

	2008-10-17	1.3.2.2	/cdg
	- Added .ini variable for DotNet folder name

	2008-11-04	1.3.2.3	/cdg
	- Added UniqueLayouts parameter to [Visual CRM]

	2008-11-06	1.3.2.4	/cdg
	- Added Policy Profile for specific Computer
	i.e if {machine} Policy Profile exists merge contents to user profile
	after all others. Designed to handle MPC program location change on
	terminal server / citrix server at PCI.

	2008-11-06	1.3.2.5	/cdg
	- Added purge of VMBTS log files

	2008-11-24	1.3.3.0	/cdg
	- Modified for manual operation at client logging on locally.
	Pimarily created for S&W Metals, requires adding VELogon.exe to
	either All Users\Startup or the registry as a local GPO using
	GPEDIT.msc to create a local logon script call.

	2008-11-24	1.3.3.1	/cdg
	- Corrected reeoe in setting VQPath for CRM

	2008-12-30	1.3.4.0	/cdg
	- Restore functionality to capture existing local .ini and .vms files
	- Added MoveOldProfile and OldProfileExt parameters

	2008-12-30	1.3.5.0	/cdg
	- Added Active Directory Group Profile

	2008-12-31	1.3.5.1	/cdg
	- Move ADGroup Profiles to ADGroup sub folder

	2008-12-31	1.3.5.2	/cdg
	- Add [Visual.Net] section and parameters
	(supersedes [DotNet] section)

	2009-01-13	1.4.0.0	/cdg
	- Added GUI display

	2009-01-17	1.4.0.1	/cdg
	- Fixes typo in DotNetInstallPath

	2009-01-21	1.4.0.2	/cdg
	- Fixed MergeINIFiles function where section exists but has no parameters

	2009-01-24	1.4.0.3	/cdg
	- Added Pre650VMFG switch to support pre 6.5.x Visual Registry
	- Added fix and COM error handling for AD Read

	2009-01-25	1.4.0.4	/cdg
	- Copyright year update.
	- Removed extra "." from OldProfileExt search string.

	2009-01-26	1.4.0.5	/cdg
	- Fixed VMFG and VQ test is Pre650.
	- Changed parameter switch from Pre650VMFG to Pre650visual

	2009-01-27	1.4.0.6	/cdg
	- Fixed Policy Profile update of local_pc_logons

	2009-08-03	1.4.1.0	/cdg
	- Added preliminary updates from 2009-01-29 - 1.4.0.7:
	.	- Corrected logic in VMFG_Installed, VQ_Installed and CRM_Installed functions
	.	- Expanded test of registry keys in HKLM VMFG_Installed function

	2009-08-03	1.4.1.1	/cdg
	- Added HKCU registry settings to associate .VMX files with Visual Manufacturing
	.	Note: Requires VEStartup updates to setup HKCR and HKLM registry settings

	2009-08-03	1.4.1.2	/cdg
	- Added HKCU registry settings to associate .VFX files with Visual Financials
	- Added HKCU registry settings to associate .VQX files with Visual Quality
	.	Note: File associations require VEStartup updates to setup HKCU registry settings

	2009-08-03	1.4.1.3	/cdg
	- Changed VM/VQ app installed test

	2009-08-04	1.4.1.4	/cdg
	- Added email configuration for Workflow etc... to both VM and VQ
	.	(VQ supports Workflow as of 6.5.3)

	2009-08-04	1.4.1.5	/cdg
	- Fixed test of results of GetIniFlag(), {key_not_defined} result, test
	.	must be converted to a string

	2009-08-05	1.4.1.6	/cdg
	- Added domain {UserID}@{Computer Name} as Display Name in VMEmail & VQEmail

	2009-08-06	1.4.1.7	/cdg
	- Relocated GUI Global variables to Definitions section at top

	2009-11-16	1.4.2.x /cdg
	- Substitute dots (.) in profile paths with underscore (_) characters
	- Replace '.' with '_' on existing profile paths
	- Migrate from '.' path to '_' path and purge old '.' path

	2010-02-26	1.4.3.x	/cdg
	x Fixup handling of dual release, pre 6.5.x and post 6.5.x Visual
	-	- Added Pre65xProfileFolder

	2010-03-25	1.4.4.x	/cdg
	- Added manipulation of read-only attribute in profile copy
	-	- Added ProfileReadOnlyCopy parameter
	-	- Added ProfileReadOnlyExt parameter

	2010-07-06	1.4.5.x	/cdg
	- Decompiled updated of 1.4.5.22 made for USN which added currency digits
	-	- Management

	2010-08-12	1.4.6.x	/cdg
	- Added CRM options for "Use Outlook Form and Email grid" and
	-   - "Don't Display Preview Pane"
	- Fixed CRM last user overwrite issue (testing HKLM instead of HKCU)

	2010-08-16	1.4.7.x	/cdg
	- Added replacement of dots (.) in username with underscore (_) in profiles.
	- Changed GetIniFlag function to return "True" / "False" / "{key_not_defined}"
	-	- instead of True / False / "{key_not_defined}" which affected all areas
	-	- using the function as a test.

	2011-02-23	1.4.8.x	/cdg
	- Added Delete Lilly Software Key

	2011-04-26	1.5.0.x	/cdg
	- Added CRM Online Books Directory setting from HKLM InstallDir
	- Added CRM .VMX configuration
	- Added CRM setting "Launch VM using .vmx file" option

	2011-08-31	1.5.1.x /cdg
	- Corrected VQ detect logic (added VQ_Installed() test)

	2012-08-24	1.6.0.x	/cdg
	- Added push command line option to force VEProfiles Update

	2012-12-07	1.7.0.x		/cdg
	- Updated VE71x/VQ71x signatures for test of app present (i.e. "Infor 10 ERP Express...")
	- Added new ToolPalette registry defaults for "ChildSmallIcons" and "SmallIcons"
	- Updated CopyRight for Central Islip and 2012

	2013-01-04	1.7.1.x		/cdg
	- Added retry for VElocalPath for those using mapped drives

	2013-01-05	1.7.2.x		/cdg
	- Added logic to capture initial Visual profile from VElocalPath if registry
	-	key HKCU/.../LocalDirectory does not exist or is empty

	2013-01-07	1.7.3.x		/cdg
	- Added VELocalPathRetry=0 to ignore failed VELocalPath search

	2013-02-15	1.7.4.x		/cdg
	- Fix VE 7.1.x VM_Installed logic

	2013-10-30	1.7.5.x		/cdg
	- Fix move old profile where if Local Directory not defined set it to
	-	$VELocalPath
	-	not
	-	$VELocalPath & "\VMFG"
	-	prior to move

	2013-11-01	1.7.6.x		/cdg
	- Fix VEProfiles$ user folder name when using local_pc_login
	- Change local_pc_login to Local-PC-Logins for readability
	- Added $VEProfileCommonGroupName for Customers without a domain
	-	ProfileCommonGroupName, if defined, takes the place of the domain in VEProfiles$

	2014-02-06	1.8.0.x		/cdg
	- Added capture of existing CRM Layout and Merge folders for new profiles

	2014-11-10	1.8.1.x	/cdg
	- Add optional installed clues paths
	- Move IGSKey test

	2014-11-20	1.8.2.x	/cdg
	- Fixed logic when clues are defined but not matched
	- Added setting HKCU VMFG Key "Use Local Directory=N

	2014-12-06	1.8.3.x	/cdg
	- Removed extra misplaced test for Infor HKLM Keys

	2017-03-30	1.8.5.x	/cdg
	- Fix DotNetOverwrite flag, test for "True" not "1"


#ce ----------------------------------------------------------------------------
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>


;----------------------------------------------------------------------------------------
; Definitions
;----------------------------------------------------------------------------------------

Dim $RegOverWriteFlag = 0
Dim $FileOverWrite = 1
Dim $Infor_HKCU_Base = "HKEY_CURRENT_USER\Software\Infor Global Solutions"
Dim $Infor_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\Infor Global Solutions"
Dim $Lilly_HKCU_Base = "HKEY_CURRENT_USER\Software\Lilly Software"
Dim $Lilly_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\Lilly Software"
Dim $VMInstallPath = ""
Dim $VQInstallPath = ""

Dim $oCOMError

; NOTE: DO NOT include trailing backslash ("\") character in path definitions

; GUI support settings
Dim $GUI_Delay = 250 ; Slow down GUI message display (ms, 1000 = 1 sec)
Dim $ScriptName = StringLeft(@ScriptName, StringLen(@ScriptName) - 4)
Dim $vScriptDesc = "Visual Enterprise Client Logon Script"
Dim $Banner_Title = "Visual Enterprise" & @CRLF & "Client Administration Scripts"
Dim $ScriptFile_ini = @ScriptDir & "\" & $ScriptName & ".ini"
Dim $ScriptTempDir = @TempDir & "\$" & $ScriptName & "$"

; (1.4.1.7)
Global $guiMsgBox
Global $guiLogo
Global $guiMain

Dim $ScriptFile_ini = @ScriptDir & "\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 3) & "ini"

If FileExists($ScriptFile_ini) Then

	; Collect Alternate Local Installation Clues (+1.8.3.x)
	;-------------------------------------------------------
	$VELogonVMInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VELogon", "VMInstalledClue", "{key_not_defined}"))
	$VELogonVQInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VELogon", "VQInstalledClue", "{key_not_defined}"))
	$VELogonCRMInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VELogon", "CRMInstalledClue", "{key_not_defined}"))
	$VELogonDotNetInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VELogon", "DotNetInstalledClue", "{key_not_defined}"))

	If $VELogonVMInstalledClue[1] = "{key_not_defined}" _
			And $VELogonVQInstalledClue[1] = "{key_not_defined}" _
			And $VELogonCRMInstalledClue[1] = "{key_not_defined}" _
			And $VELogonDotNetInstalledClue[1] = "{key_not_defined}" Then

		;----------------------------------------------------------------------------------------
		; If Infor Products do not exist Exit immediately
		;----------------------------------------------------------------------------------------
		$IGSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Infor Global Solutions", "")
		; Infor products do not exist
		If @error > 0 Then
			$LSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Lilly Software", "")
			If @error > 0 Then
				Exit
			EndIf
		EndIf
	Else

		; If VMInstalledClue is defined look for valid path else use HKLM technique
		$VMInstalledFlag = False
		If Not ($VELogonVMInstalledClue[0] = 1 And $VELogonVMInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VELogonVMInstalledClue[0]
				If FileExists($VELogonVMInstalledClue[$i]) Then
					$VMInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If VQInstalledClue is defined look for valid path else use HKLM technique
		$VQInstalledFlag = False
		If Not ($VELogonVQInstalledClue[0] = 1 And $VELogonVQInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VELogonVQInstalledClue[0]
				If FileExists($VELogonVQInstalledClue[$i]) Then
					$VQInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If CRMInstalledClue is defined look for valid path else use HKLM technique
		$CRMInstalledFlag = False
		If Not ($VELogonCRMInstalledClue[0] = 1 And $VELogonCRMInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VELogonCRMInstalledClue[0]
				If FileExists($VELogonCRMInstalledClue[$i]) Then
					$CRMInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If DotNetInstalledClue is defined look for valid path else use HKLM technique
		$DotNetInstalledFlag = False
		If Not ($VELogonDotNetInstalledClue[0] = 1 And $VELogonDotNetInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VELogonDotNetInstalledClue[0]
				If FileExists($VELogonDotNetInstalledClue[$i]) Then
					$DotNetInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		If $VMInstalledFlag = False _
				And $VQInstalledFlag = False _
				And $CRMInstalledFlag = False _
				And $DotNetInstalledFlag = False Then
			Exit
		EndIf
	EndIf


	;----------------------------------------------------------------------------------------
	; GUI Display Setup
	;----------------------------------------------------------------------------------------
	; Include GUI support files
	If FileExists($ScriptTempDir) Then FileDelete($ScriptTempDir)
	DirCreate($ScriptTempDir)
	FileInstall("SynergyPH3.JPG", $ScriptTempDir & "\")

	; Build main GUI display
	GUIBuildMain()

	; Display GUI
	GUISetState(@SW_SHOW)


	;---------------------------------------------------------------------------------------------
	; Collect parameters from VELogon.ini file
	;---------------------------------------------------------------------------------------------
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Recording script info in registry...")
	Sleep($GUI_Delay)

	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "@ScriptDir", "REG_SZ", @ScriptDir)
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "@ScriptName", "REG_SZ", @ScriptName)
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "RunTime", "REG_SZ", @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC)
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "")
	$ScriptVersion = FileGetVersion(@ScriptFullPath)
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ScriptVersion", "REG_SZ", $ScriptVersion)
	$IniVersion = FileGetTime($ScriptFile_ini, 0, 1)
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "IniVersion", "REG_SZ", $IniVersion)


	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Reading " & $ScriptName & ".ini script configuration file...")
	Sleep($GUI_Delay)

	;-------------------------------
	; Read [VELogon] Section
	;-------------------------------
	$VELogonVELocalPath = CleanPath(GetIniArray($ScriptFile_ini, "VELogon", "VELocalPath", ""))

	; Added in 1.7.1.x
	;	Retry delay for VELocalPath
	$VELogonVELocalPathTimeout = IniRead($ScriptFile_ini, "VELogon", "VELocalPathTimeout", "5")
	$VELogonVELocalPathRetry = IniRead($ScriptFile_ini, "VELogon", "VELocalPathRetry", "1")

	$VELogonVEProfilesPath = IniRead($ScriptFile_ini, "VELogon", "VEProfilesPath", "")
	; Added in 1.3.1.2
	;	Location to record BuildInformation keys. Default is the same location as the script.ini
	;	as {Script_ini_Name}.BILog.ini. There must be only one file enterprise-wide.
	$VELogonVEBuildInformationLogFile = IniRead($ScriptFile_ini, "VELogon", "VEBuildInformationLogFile", StringLeft($ScriptFile_ini, StringLen($ScriptFile_ini) - 3) & "BILog.ini")

	; Flags to disable configurations. Default for all is TRUE.
	$VEVELogonConfigVMFlag = GetIniFlag($ScriptFile_ini, "VELogon", "ConfigVM", "True")
	$VEVELogonConfigVQFlag = GetIniFlag($ScriptFile_ini, "VELogon", "ConfigVQ", "True")
	$VEVELogonConfigCRMFlag = GetIniFlag($ScriptFile_ini, "VELogon", "ConfigCRM", "True")
	$VEVELogonConfigDotNetFlag = GetIniFlag($ScriptFile_ini, "VELogon", "ConfigDotNet", "True")
	$VEVELogonMoveOldProfileFlag = GetIniFlag($ScriptFile_ini, "VELogon", "MoveOldProfile", "False")
	$VEVELogonDeleteLillyRegKeysFlag = GetIniFlag($ScriptFile_ini, "VELogon", "DeleteLillyRegKeys", "False")

	; Delete Lilly Software HKCU key if DeleteLillyRegKeys = True
	If $VEVELogonDeleteLillyRegKeysFlag = "True" Then
		RegDelete($Lilly_HKCU_Base)
	EndIf


	; Setup for pre 6.5.x Visual registry
	$VEVELogonPre650VisualFlag = GetIniFlag($ScriptFile_ini, "VELogon", "Pre650Visual", "False")
	If $VEVELogonPre650VisualFlag = "True" Then
		$Reg_HKLM_Base = $Lilly_HKLM_Base
		$Reg_HKCU_Base = $Lilly_HKCU_Base
	Else
		$Reg_HKLM_Base = $Infor_HKLM_Base
		$Reg_HKCU_Base = $Infor_HKCU_Base
	EndIf

	; List of user profile file extensions to move if MoveOldProfile is set
	$VEVELogonOldProfileExt = CleanExt(GetIniArray($ScriptFile_ini, "VELogon", "OldProfileExt", ".ini, .vms"))

	; Maintain VMBTS logs purging files older than x days, default = 30 days, 0 = no purge
	$VEVELogonPurgeVMBTSLogDays = IniRead($ScriptFile_ini, "VELogon", "PurgeVMBTSLogDays", "30")

	; Active Directory Group Memberships to test for (added in 1.3.5.0 /cdg)
	$VEVELogonADGroupProfiles = _StringSplitTrim(IniRead($ScriptFile_ini, "VELogon", "ADGroupProfiles", ""), ",")

	; GUI Delay Overide default = 250 (1/4 second) valid range = 0 to 10000
	$VEVELogonGUIDelay = IniRead($ScriptFile_ini, "VELogon", "VELogonGUIDelay", "250")
	If $VEVELogonGUIDelay >= 0 And $VEVELogonGUIDelay <= 10000 Then
		$GUI_Delay = $VEVELogonGUIDelay
	EndIf

	$VELogonPre65xProfileFolder = IniRead($ScriptFile_ini, "VELogon", "Pre65xProfileFolder", "")

	; Profile file copy with read-only set on target (+1.4.4.x)
	; 	ProfileReadOnlyCopy=0  -  OS default. (default)
	; 	ProfileReadOnlyCopy=1  -  Overwrite read-only destination files.
	; 	ProfileReadOnlyCopy=2  -  Preserve destination read-only attribute.
	; 	ProfileReadOnlyCopy=3  -  Force destination read-only attribute.
	$VEVELogonProfileReadOnlyCopyFlag = IniRead($ScriptFile_ini, "VELogon", "ProfileReadOnlyCopy", 0)
	$VEVELogonProfileReadOnlyCopyExt = CleanExt(GetIniArray($ScriptFile_ini, "VELogon", "ProfileReadOnlyCopyExt", ".vms"))

	; Set number of digits used for currency (+1.4.5.x)
	$VEVELogonCurrencyDigits = StringStripWS(IniRead($ScriptFile_ini, "VELogon", "CurrencyDigits", ""), 3)
	If $VEVELogonCurrencyDigits <> "" Then
		If StringLen($VEVELogonCurrencyDigits) = 1 Then
			If Not Asc($VEVELogonCurrencyDigits) >= Asc("0") And Asc($VEVELogonCurrencyDigits) <= Asc("9") Then
				MsgBox(0, "VELogon - Error: CurrencyDigits Out of Range", $ScriptFile_ini & @CRLF & "The value specified for CurrencyDigits must be between 0 and 9" & @CRLF & @CRLF & "Please report this error to your Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "CurrencyDigits Out of Range")
				Exit
			EndIf
		Else
			MsgBox(0, "VELogon - Error: CurrencyDigits Out of Range", $ScriptFile_ini & @CRLF & "The value specified for CurrencyDigits must be between 0 and 9" & @CRLF & @CRLF & "Please report this error to your Administrator")
			RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "CurrencyDigits Out of Range")
			Exit
		EndIf
		RegWrite("HKEY_CURRENT_USER\Control Panel\International", "iCurrDigits", "REG_SZ", $VEVELogonCurrencyDigits)
	EndIf

	; +1.7.6.x
	; For environments without an Active Directory domain, specify the name to use in its place
	;   in VEProfiles$ folder names. Leave undefines to use triditional user LogonDomain
	$VEProfileCommonGroupName = IniRead($ScriptFile_ini, "VELogon", "ProfileCommonGroupName", "")


	;-------------------------------
	; Read [Visual CRM] Section
	;-------------------------------
	$CRMDBTypes = _StringSplitTrim(IniRead($ScriptFile_ini, "Visual CRM", "DatabaseTypes", "SQL Server"), ",")
	$CRMOLEDBProviders = _StringSplitTrim(IniRead($ScriptFile_ini, "Visual CRM", "OLEDBProviders", "SQLOLEDB\Microsoft OLE DB Provider for SQL Server"), ",")
	$CRMDatabases = _StringSplitTrim(IniRead($ScriptFile_ini, "Visual CRM", "Databases", "{key_not_defined}"), ",")
	$CRMServerNames = _StringSplitTrim(IniRead($ScriptFile_ini, "Visual CRM", "ServerNames", "{key_not_defined}"), ",")
	$CRMDBType = IniRead($ScriptFile_ini, "Visual CRM", "DatabaseType", "")
	$CRMDatabase = IniRead($ScriptFile_ini, "Visual CRM", "Database", "")
	$CRMOLEDBProvider = IniRead($ScriptFile_ini, "Visual CRM", "OLEDBProvider", "")
	$CRMServer = IniRead($ScriptFile_ini, "Visual CRM", "Server", "{key_not_defined}")
	$CRMUniqueLayouts = GetIniFlag($ScriptFile_ini, "Visual CRM", "UniqueLayouts", "{key_not_defined}")
	;+1.4.6.x
	$CRMEmployOutlookForm = GetIniFlag($ScriptFile_ini, "Visual CRM", "EmployOutlookForm", "{key_not_defined}")
	$CRMHidePreviewPane = GetIniFlag($ScriptFile_ini, "Visual CRM", "HidePreviewPane", "{key_not_defined}")

	;+1.5.0.x
	$CRMUseVMXforVMLaunch = GetIniFlag($ScriptFile_ini, "Visual CRM", "UseVMXforVMLaunch", "{key_not_defined}")

	;-------------------------------
	; Read [DotNet] Section
	;-------------------------------
	$DotNetAutoUpdateURL = IniRead($ScriptFile_ini, "DotNet", "AutoUpdateURL", "")
	$DotNetReportsPathBase = IniRead($ScriptFile_ini, "DotNet", "DotNetReportsPathBase", "")
	$DotNetInstallList = CleanPath(GetIniArray($ScriptFile_ini, "DotNet", "DotNetInstallPath", ""))

	;+1.3.5.2{
	;-------------------------------
	; Read [Visual.Net] Section
	;-------------------------------
	$DotNetAutoUpdateURL = IniRead($ScriptFile_ini, "Visual.Net", "AutoUpdateURL", "")
	$DotNetReportsPathBase = IniRead($ScriptFile_ini, "Visual.Net", "DotNetReportsPathBase", "")
	$DotNetInstallList = CleanPath(GetIniArray($ScriptFile_ini, "Visual.Net", "DotNetInstallPath", "DotNet"))

	$DotNetOverwriteFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "DotNetOverwrite", "True")
	$DotNetlsa_ShowErrorMsgBoxFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "lsa_ShowErrorMsgBox", "True")
	$DotNetlsa_SaveSettingsOnExitFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "lsa_SaveSettingsOnExit", "False")
	$DotNetlsa_LogErrorsFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "lsa_LogErrorsFlag", "False")
	$DotNetlsa_InsertOnEndOfGridFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "lsa_InsertOnEndOfGridFlag", "True")
	$DotNetlsa_ClearOnSaveFlag = GetIniFlag($ScriptFile_ini, "Visual.Net", "lsa_ClearOnSave", "False")
	;+1.3.5.2}


	;-------------------------------
	; Read [Visual Mfg] Section
	;-------------------------------
	; Added in 1.3.1.2
	;---------------------------------------------------------------------------------------------
	;	Search list for application location. Set first location where vm.exe is found.
	If $VEVELogonConfigVMFlag = "True" Then
		$VMInstallPathList = CleanPath(GetIniArray($ScriptFile_ini, "Visual Mfg", "VMInstallPath", ""))
		If $VMInstallPathList[1] <> "" Then
			For $i = 1 To $VMInstallPathList[0]
				If FileExists($VMInstallPathList[$i] & "\vm.exe") Then
					$VMInstallPath = $VMInstallPathList[$i]
					ExitLoop
				EndIf
			Next
			If $VMInstallPath = "" Then
				MsgBox(0, "VELogon - Error: Cannot find VM.EXE", _
						$ScriptFile_ini & @CRLF & _
						"VM.EXE not found in any VMInstallPath locations" & @CRLF & @CRLF & _
						"Please report this error to your Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VM.EXE")
				Exit
			EndIf
		EndIf
		;---------------------------------------------------------------------------------------------
		$VMOnlineBooksInstallPath = CleanPath(GetIniArray($ScriptFile_ini, "Visual Mfg", "VMOnlineBooksInstallPath", ""))

	EndIf


	;-----------------------------------
	; Read [VMEmail] Section (+1.4.1.4)
	;-----------------------------------
	$VMEmailAutoDiscover = GetIniFlag($ScriptFile_ini, "VMEmail", "AutoDiscover", "{key_not_defined}")
	$VMEmailSMTPServer = IniRead($ScriptFile_ini, "VMEmail", "SMTPServer", "{key_not_defined}")
	$VMEmailSMTPEmailAddress = IniRead($ScriptFile_ini, "VMEmail", "SMTPEmailAddress", "{key_not_defined}")
	$VMEmailExchangeUser = IniRead($ScriptFile_ini, "VMEmail", "ExchangeUser", "{key_not_defined}")
	$VMEmailDisplayName = IniRead($ScriptFile_ini, "VMEmail", "DisplayName", "{key_not_defined}")


	;-------------------------------
	; Read [Visual Quality] Section
	;-------------------------------
	; Added in 1.3.1.2
	;---------------------------------------------------------------------------------------------
	;	Search list for application location. Set first location where vm.exe is found.
	If $VEVELogonConfigVQFlag = "True" Then
		If VQ_Installed() Then ;+1.5.1.x
			$VQInstallPathList = CleanPath(GetIniArray($ScriptFile_ini, "Visual Quality", "VQInstallPath", ""))
			If $VQInstallPathList[1] <> "" Then
				For $i = 1 To $VQInstallPathList[0]
					If FileExists($VQInstallPathList[$i] & "\vq.exe") Then
						$VQInstallPath = $VQInstallPathList[$i]
						ExitLoop
					EndIf
				Next
				If $VQInstallPath = "" Then
					MsgBox(0, "VELogon - Error: Cannot find VQ.EXE", _
							$ScriptFile_ini & @CRLF & _
							$VQInstallPathList[0] & @CRLF & _
							$VQInstallPathList[1] & @CRLF & _
							"VQ.EXE not found in any VQInstallPath locations" & @CRLF & @CRLF & _
							"Please report this error to your Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VQ.EXE")
					Exit
				EndIf
			EndIf ;+1.5.1.x
		EndIf
		;---------------------------------------------------------------------------------------------
		$VQOnlineBooksInstallPath = CleanPath(GetIniArray($ScriptFile_ini, "Visual Quality", "VQOnlineBooksInstallPath", ""))
	EndIf

	;-----------------------------------
	; Read [VQEmail] Section (+1.4.1.4)
	;-----------------------------------
	$VQEmailUseVMSettings = IniRead($ScriptFile_ini, "VQEmail", "UseVMSettings", "{key_not_defined}")
	$VQEmailAutoDiscover = IniRead($ScriptFile_ini, "VQEmail", "AutoDiscover", "{key_not_defined}")
	$VQEmailSMTPServer = IniRead($ScriptFile_ini, "VQEmail", "SMTPServer", "{key_not_defined}")
	$VQEmailSMTPEmailAddress = IniRead($ScriptFile_ini, "VQEmail", "SMTPEmailAddress", "{key_not_defined}")
	$VQEmailExchangeUser = IniRead($ScriptFile_ini, "VQEmail", "ExchangeUser", "{key_not_defined}")
	$VQEmailDisplayName = IniRead($ScriptFile_ini, "VQEmail", "DisplayName", "{key_not_defined}")



Else
	MsgBox(0, "VELogon - Error: Missing VELogon.ini", _
			$ScriptFile_ini & @CRLF & _
			"VELogon configuration file missing or inaccessable" & @CRLF & @CRLF & _
			"Please report this error to your Administrator")
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Missing VELogon.ini")
	Exit
EndIf

;---------------------------------------------------------------------------------------------
; Validate VELogon.ini settings
;---------------------------------------------------------------------------------------------

; Locate VELocalPath (i.e. c:\Visual)
; Added Timeout and Retry in 1.7.1.x
$VELocalPath = ""
$VELocalPathRetryCount = 0
While True
	For $p = 1 To $VELogonVELocalPath[0]
		If FileExists($VELogonVELocalPath[$p]) Then
			$VELocalPath = $VELogonVELocalPath[$p]
			ExitLoop
		EndIf
	Next
	If $VELocalPath = "" Then
		;+1.7.3.x{
		; if VELocalPathRetry = 0 then force to last in the VELocalPath list
		If $VELogonVELocalPathRetry = 0 Then
			$VELocalPath = $VELogonVELocalPath[$VELogonVELocalPath[0]]
			ExitLoop
		EndIf
		;}+1.7.3.x
		If $VELocalPathRetryCount < $VELogonVELocalPathRetry Then
			$VELocalPathRetryCount += 1
			Sleep($VELogonVELocalPathTimeout * 1000)
		Else
			ExitLoop
		EndIf
	Else
		ExitLoop
	EndIf
WEnd
RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "VELocalPathRetryCount", "REG_SZ", $VELocalPathRetryCount)
RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "VELocalPathRetry", "REG_SZ", $VELogonVELocalPathRetry)
RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "VELocalPathTimeout", "REG_SZ", $VELogonVELocalPathTimeout)

If $VELocalPath = "" Then
	MsgBox(0, "VELogon - Error: Cannot find VE Application local install path", _
			$ScriptFile_ini & @CRLF & _
			"No valid local application path found in any VELocalPath locations" & @CRLF & @CRLF & _
			"Please report this error to your Administrator")
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VELocalPath")
	Exit
EndIf

; Locate Visual.Net install path (i.e. c:\Visual\Visual.Net)
$DotNetInstallPath = ""
For $p = 1 To $DotNetInstallList[0]
	If FileExists($DotNetInstallList[$p]) Then
		$DotNetInstallPath = $DotNetInstallList[$p]
		ExitLoop
	EndIf
Next


;$IGSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Infor Global Solutions", "")
;; Infor products do not exist
;If @error > 0 Then
;	$LSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Lilly Software", "")
;	If @error > 0 Then
;		Exit
;	EndIf
;EndIf


;---------------------------------------------------------------------------------------------
; Seed new user profile
;---------------------------------------------------------------------------------------------
;$RegTmp = RegRead($Reg_HKLM_Base, "")
;If @error <> 1 Then ;Infor Products are installed
If True Then ; With the new InstalledClues previous tests will determin
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Building user profile folder hierarchy ...")
	Sleep($GUI_Delay)

	; Added in 1.7.6.x	/cdg {
	If $VEProfileCommonGroupName <> "" Then
		$VEProfileDomain = StringReplace($VEProfileCommonGroupName, ".", "_")
	Else

		; Added in 1.3.3.0	/cdg {
		If @LogonDNSDomain <> "" Then
			$VEProfileDomain = StringLower(StringReplace(@LogonDNSDomain, ".", "_"))
		ElseIf StringLower(@ComputerName) = StringLower(@LogonDomain) Then
			$VEProfileDomain = "Local-PC-Logons"
		Else
			MsgBox(0, @ScriptName & " - Error", "Unable to determine the user's logon domain" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator")
			RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to determine the user's logon domain")
			Exit
		EndIf
		; }
	EndIf
	; }


	;---------------------------------------------------------------------------------------------
	; Build profile folder hierarchy if missing
	;---------------------------------------------------------------------------------------------
	; Added in 1.3.1.1
	;	Initially create all folders necessary for VEProfile:
	;		\\server\VEProfiles$\domain_dns\
	;		\\server\VEProfiles$\domain_dns\Default Profile
	;		\\server\VEProfiles$\domain_dns\Default Profile\VMFG
	;		\\server\VEProfiles$\domain_dns\Default Profile\VMFG641		(Pre65xProfileFolder +1.4.3.x)
	;		\\server\VEProfiles$\domain_dns\Default Profile\VQ

	;		\\server\VEProfiles$\domain_dns\Default Profile\CRM\Layout
	;		\\server\VEProfiles$\domain_dns\Default Profile\CRM\Merge Docs
	;		\\server\VEProfiles$\domain_dns\Default Profile\CRM\Merge Temp
	;		\\server\VEProfiles$\domain_dns\Policy Profile
	;		\\server\VEProfiles$\domain_dns\Policy Profile\VMFG
	;		\\server\VEProfiles$\domain_dns\Policy Profile\VMFG641		(Pre65xProfileFolder +1.4.3.x)
	;		\\server\VEProfiles$\domain_dns\Policy Profile\VQ
	;		\\server\VEProfiles$\domain_dns\Policy Profile\CRM\Layout
	;		\\server\VEProfiles$\domain_dns\Policy Profile\CRM\Merge Docs
	;		\\server\VEProfiles$\domain_dns\Policy Profile\CRM\Merge Temp
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\VMFG
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\VMFG641		(Pre65xProfileFolder +1.4.3.x)
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\VQ
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\CRM\Layout
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\CRM\Merge Docs
	;		\\server\VEProfiles$\domain_dns\ADGroup Profiles\{ADGroupName} Policy Profile\CRM\Merge Temp
	;---------------------------------------------------------------------------------------------
	; Create profile base dns folder if missing
	If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain) Then
		If FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain) Then
			DirCopy($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\Default Profile", $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\Default Profile")
			DirCopy($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\Policy Profile", $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\Policy Profile")
			DirCopy($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\ADGroup Profiles", $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\ADGroup Profiles")
		Else
			$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain)
			If Not $DCOK Then
				MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create profile base dns folder")
				Exit
			EndIf
		EndIf

	EndIf


	; If legacy folder (with (.) in name, if no user profiles remain remove top directory
	If FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain) Then
		$hProfile = FileFindFirstFile($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\*")
		If $hProfile = -1 Then
			; directory is empty so remove it
			DirRemove($VELogonVEProfilesPath & "\" & $VEProfileDomain, 1)
		Else
			While 1
				$PSFile = FileFindNextFile($hProfile)
				If @error Then
					DirRemove($VELogonVEProfilesPath & "\" & $VEProfileDomain, 1)
					ExitLoop
				Else
					If StringInStr("|Default Profile|Policy Profile|ADGroup Profiles|", $PSFile) = 0 Then
						ExitLoop
					EndIf
				EndIf
			WEnd
		EndIf
		FileClose($hProfile)
	EndIf


	;+1.3.5.0{
	$PTmp = "Default Profile, Policy Profile"
	For $i = 1 To $VEVELogonADGroupProfiles[0]
		If $VEVELogonADGroupProfiles[$i] <> "" Then
			$PTmp &= ", ADGroup Profiles\{" & $VEVELogonADGroupProfiles[$i] & "} Policy Profile"
		EndIf

	Next
	$ProfilesFolders = _StringSplitTrim($PTmp, ",")


	For $i = 1 To $ProfilesFolders[0]

		; Create profile base hierarchy if missing
		If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i]) Then
			$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i])
			If Not $DCOK Then
				MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile folder")
				Exit
			EndIf
		EndIf

		; Create VMFG folder if missing
		If VMFG_Installed() Then
			If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\VMFG") Then
				$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\VMFG")
				If Not $DCOK Then
					MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
							"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
							"Please report this error to your System Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\VMFG folder")
					Exit
				EndIf
			EndIf
		EndIf

		; Create VMFG641 folder if missing	(Pre65xProfileFolder +1.4.3.x)
		If VMFG_Installed() Then
			If $VELogonPre65xProfileFolder <> "" Then
				If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\" & $VELogonPre65xProfileFolder) Then
					$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\" & $VELogonPre65xProfileFolder)
					If Not $DCOK Then
						MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
								"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
								"Please report this error to your System Administrator")
						RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\" & $VELogonPre65xProfileFolder)
						Exit
					EndIf
				EndIf
			EndIf
		EndIf

		; Create VQ folder if missing
		If VQ_Installed() Then
			If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\VQ") Then
				$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\VQ")
				If Not $DCOK Then
					MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
							"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
							"Please report this error to your System Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\VQ folder")
					Exit
				EndIf
			EndIf
		EndIf

		; Create CRM folders if missing
		If CRM_Installed() Then
			If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Layout") Then
				$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Layout")
				If Not $DCOK Then
					MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
							"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
							"Please report this error to your System Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\CRM\Layout")
					Exit
				EndIf
			EndIf
			If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Merge Docs") Then
				$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Merge Docs")
				If Not $DCOK Then
					MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
							"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
							"Please report this error to your System Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\CRM\Merge Docs")
					Exit
				EndIf
			EndIf
			If Not FileExists($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Merge Temp") Then
				$DCOK = DirCreate($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $ProfilesFolders[$i] & "\CRM\Merge Temp")
				If Not $DCOK Then
					MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
							"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
							"Please report this error to your System Administrator")
					RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot create Default Profile\CRM\Merge Temp")
					Exit
				EndIf
			EndIf

		EndIf

	Next
	;}+1.3.5.0
	;---------------------------------------------------------------------------------------------

	;+1.3.3.0{
	If $VEProfileDomain = StringReplace(@LogonDNSDomain, ".", "_") Then
		$VEProfile_Path = StringLower($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & StringReplace(@UserName, ".", "_") & "_" & $VEProfileDomain)
	Else
		$VEProfile_Path = StringLower($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & StringReplace(@UserName, ".", "_") & "_" & $VEProfileDomain)
	EndIf
	;}+1.3.3.0

	If Not FileExists($VEProfile_Path) Then

		If FileExists(StringReplace($VEProfile_Path, "_", ".")) Then

			$DCOK = DirCopy(StringReplace($VEProfile_Path, "_", "."), $VEProfile_Path)
			If $DCOK Then
				DirRemove(StringReplace($VEProfile_Path, "_", "."), 1)
			EndIf

		Else

			$DCOK = DirCopy($VELogonVEProfilesPath & "\" & $VEProfileDomain & "\Default Profile", $VEProfile_Path)
			If Not $DCOK Then
				MsgBox(0, @ScriptName & " - Error", "Unable to create Visual Enterprise Profile" & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot copy Default Profile to user profile")
				Exit
			EndIf

			;+1.3.4.0{
			If $VEVELogonMoveOldProfileFlag = "True" Then
				$OldProfilePath = RegRead($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "Local Directory")
				;+1.7.2.x{
				; if the Local Directory key does not exist or does not exist, copy from VELocalPath
				If $OldProfilePath = "" Then
					$OldProfilePath = $VELocalPath
				EndIf

				;}1.7.2.x
				If FileExists($OldProfilePath) Then
					For $i = 1 To $VEVELogonOldProfileExt[0]
						$FCOK = FileCopy($OldProfilePath & "\*" & $VEVELogonOldProfileExt[$i], $VEProfile_Path & "\VMFG", 9)
						; We don't care if the copy fails, it will fail on no wildcard hits
					Next
				EndIf
				;+1.4.3.x{
				If $VELogonPre65xProfileFolder <> "" Then
					$Old65xProfilePath = RegRead($Lilly_HKCU_Base & "\Visual Manufacturing\Configuration", "Local Directory")
					If FileExists($Old65xProfilePath) Then
						For $i = 1 To $VEVELogonOldProfileExt[0]
							$FCOK = FileCopy($Old65xProfilePath & "\*" & $VEVELogonOldProfileExt[$i], $VEProfile_Path & "\" & $VELogonPre65xProfileFolder, 9)
							; We don't care if the copy fails, it will fail on no wildcard hits
						Next
					EndIf
				EndIf
				;}+1.4.3.x
			EndIf
			;}+1.3.4.0

		EndIf
	Else
		If FileExists(StringReplace($VEProfile_Path, "_", ".")) Then
			DirRemove(StringReplace($VEProfile_Path, "_", "."), 1)
		EndIf

	EndIf

Else

	; No Infor products installed
	Exit
EndIf


If VMFG_Installed() Or VQ_Installed() Then ; (+1.4.1.3)

	;---------------------------------------------------------------------------------------------
	; Manage BuildInformation registry key
	;---------------------------------------------------------------------------------------------
	; Added in 1.3.1.1
	;	The BuildInformation registry key contains the source for the hashed clients Machine ID
	;	in the LOGINS table in the Visual databases associated with the issueing of licenses.
	;	This key must be unique throughout the domain in order for Visual licensing to work
	;	properly. The key is initially created by the first sucessful login to Visual by a
	;	particular domain user from a particular machine. In a multi-user environment such as
	;	Terminal Server or Citrix, this key is often duplicated from the key generated during
	;	install creating the potential for L101 licensing errors in Visual.
	;
	;	This section of the program will maintain a master list of BuildInformation keys in the
	;	.ini file associated with this script. If a duplicate key is found in the .ini file it
	;	will reset the key in the registry removing the potential for duplication.
	;---------------------------------------------------------------------------------------------

	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Configuring Visual Manufacturing user registry settings...")
	Sleep($GUI_Delay)

	$BuildInformation = RegRead("HKEY_CURRENT_USER\Software\Identification\Other\BuildInformation", "")

	If $BuildInformation <> "" Then

		; If BuildInformation Reg value exists check for duplicate elsewhere
		$BIiniTemp = IniRead($VELogonVEBuildInformationLogFile, "BuildInformation", $BuildInformation, "")
		If $BIiniTemp = "" Then

			; If BuildInformation key has not been previously recorded in .ini file, do so
			$IWOK = IniWrite($VELogonVEBuildInformationLogFile, "BuildInformation", $BuildInformation, StringLower(@UserName & "." & $VEProfileDomain & "@" & @ComputerName))
			If $IWOK <> 1 Then
				MsgBox(0, @ScriptName & " - Error", "Unable to BILog.ini file" & @CRLF & _
						$VELogonVEBuildInformationLogFile & @CRLF & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator")
				RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot write to ScriptFile.ini file")
				Exit
			EndIf


		Else
			; If BuildInformation key was previously recorded in .ini file, was it from this user/machine
			If $BIiniTemp <> StringLower(@UserName & "." & $VEProfileDomain & "@" & @ComputerName) Then

				; BuildInformation key in .ini belongs to another user/machine so clear the one on this user/machine
				; Visual will then build a new unique key when next launched
				RegWrite("HKEY_CURRENT_USER\Software\Identification\Other\BuildInformation", "", "REG_SZ", "")
			EndIf
		EndIf
	EndIf



	;---------------------------------------------------------------------------------------------
	; Initialize Visual Manufacturing user configuration
	;---------------------------------------------------------------------------------------------
	If VMFG_Installed() Then ;Visual Manufacturing is installed

		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Configuring Visual Manufacturing user registry settings...")
		Sleep($GUI_Delay)

		; Set Visual Mfg registry settings
		;----------------------------------
		RegWrite($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "Local Directory", "REG_SZ", $VEProfile_Path & "\VMFG")
		RegWrite($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "InstallDirectory", "REG_SZ", $VMInstallPath)
		RegWrite($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "Use Local Directory", "REG_SZ", "N")
		For $i = 1 To $VMOnlineBooksInstallPath[0]
			If FileExists($VMOnlineBooksInstallPath[$i]) Then
				RegWrite($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "OnlineBooksInstallPath", "REG_SZ", $VMOnlineBooksInstallPath[$i])
			EndIf
		Next
		;+1.4.3.x{
		If $VELogonPre65xProfileFolder <> "" Then
			RegWrite($Lilly_HKCU_Base & "\Visual Manufacturing\Configuration", "Local Directory", "REG_SZ", $VEProfile_Path & "\" & $VELogonPre65xProfileFolder)
		EndIf
		;}1.4.3.x

		; Create, but not overwrite these keys unless overwrite flag set
		$DefaultKey = RegRead($Reg_HKCU_Base & "\Visual Manufacturing\ToolPalette", "")
		If @error = 1 Or $RegOverWriteFlag = 1 Then
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Left", "REG_DWORD", "0x00000064")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Top", "REG_DWORD", "0x00000064")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Right", "REG_DWORD", "0x00000140")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Bottom", "REG_DWORD", "0x00000190")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "AlwaysOnTop", "REG_DWORD", "0x00000008")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Visible", "REG_DWORD", "0x00000000")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Minimized", "REG_DWORD", "0x00000000")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "ChildSmallIcons", "REG_DWORD", "0x00000001")
			RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "SmallIcons", "REG_DWORD", "0x00000001")
		EndIf
		;		;+1.4.3.x{
		;		$DefaultKey = RegRead($Lilly_HKCU_Base & "\Visual Manufacturing\ToolPalette", "")
		;		If @error = 1 Or $RegOverWriteFlag = 1 Then
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Left", "REG_DWORD", "0x00000064")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Top", "REG_DWORD", "0x00000064")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Right", "REG_DWORD", "0x00000140")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Bottom", "REG_DWORD", "0x00000190")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "AlwaysOnTop", "REG_DWORD", "0x00000008")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Visible", "REG_DWORD", "0x00000000")
		;			RegWrite($Lilly_HKCU_Base & "\VISUAL Manufacturing\ToolPalette", "Minimized", "REG_DWORD", "0x00000000")
		;		EndIf
		;		;}+1.4.3.x

		; When setting up for Pre 6.5.0, also setup Infor HKCU key for DotNet link
		If $VEVELogonPre650VisualFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Visual Manufacturing\Configuration", "InstallDirectory", "REG_SZ", $VMInstallPath)
		EndIf


		;-------------------------------------------------------------------------------
		; Setup HKCU registry settings for .VMX file association with VM.EXE (+1.4.1.1)
		;-------------------------------------------------------------------------------
		; Remove .vmx association in HKCU in favor of HKLM as setup by VEStartup
		RegDelete("HKEY_CURRENT_USER\.vmx")

		; Replace any "Open With" list
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vmx")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vmx\OpenWithList", "a", "REG_SZ", "vm.exe")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vmx\OpenWithList", "b", "REG_SZ", "NOTEPAD.EXE")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vmx\OpenWithList", "MRUList", "REG_SZ", "ab")


		;-------------------------------------------------------------------------------
		; Setup HKCU registry settings for .VFX file association with VF.EXE (+1.4.1.2)
		;-------------------------------------------------------------------------------
		; Remove .vfx association in HKCU in favor of HKLM as setup by VEStartup
		RegDelete("HKEY_CURRENT_USER\.vfx")

		; Replace any "Open With" list
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vfx")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vfx\OpenWithList", "a", "REG_SZ", "vf.exe")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vfx\OpenWithList", "b", "REG_SZ", "NOTEPAD.EXE")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vfx\OpenWithList", "MRUList", "REG_SZ", "ab")


		;------------------------------------------------------
		; Setup Visual Manufacturing Workflow Email (+1.4.1.4)
		;------------------------------------------------------

		If String($VMEmailAutoDiscover) <> "{key_not_defined}" Then
			If $VMEmailAutoDiscover Then
				RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "Auto Discover", "REG_DWORD", "0x00000001")
			Else
				RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "Auto Discover", "REG_DWORD", "0x00000000")
			EndIf
		EndIf

		If $VMEmailSMTPServer <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP Server", "REG_SZ", $VMEmailSMTPServer)
		If $VMEmailSMTPEmailAddress <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP Email Address", "REG_SZ", $VMEmailSMTPEmailAddress)
		If $VMEmailExchangeUser <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "Exchange User", "REG_SZ", $VMEmailExchangeUser)
		If $VMEmailDisplayName <> "{key_not_defined}" Then
			If $VMEmailDisplayName = "" Then
				RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP UserName", "REG_SZ", @UserName & "@" & @ComputerName)
			Else
				RegWrite($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP UserName", "REG_SZ", $VMEmailDisplayName)
			EndIf
		EndIf

	EndIf ; VMFGInstalled()


	;---------------------------------------------------------------------------------------------
	; Initialize Visual Quality user configuration
	;---------------------------------------------------------------------------------------------
	If VQ_Installed() Then ;Visual Quality is installed

		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Configuring Visual Quality user registry settings...")
		Sleep($GUI_Delay)

		; Set Visual Quality registry settings
		;--------------------------------------
		RegWrite($Reg_HKCU_Base & "\Visual Quality\Configuration", "Local Directory", "REG_SZ", $VEProfile_Path & "\VQ")
		RegWrite($Reg_HKCU_Base & "\Visual Quality\Configuration", "InstallDirectory", "REG_SZ", $VQInstallPath)
		For $i = 1 To $VQOnlineBooksInstallPath[0]
			If FileExists($VQOnlineBooksInstallPath[$i]) Then
				RegWrite($Reg_HKCU_Base & "\Visual Quality\Configuration", "OnlineBooksInstallPath", "REG_SZ", $VQOnlineBooksInstallPath[$i])
			EndIf
		Next

		; When setting up for Pre 6.5.0, also setup Infor HKCU key for DotNet link
		If $VEVELogonPre650VisualFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Visual Quality\Configuration", "InstallDirectory", "REG_SZ", $VQInstallPath)
		EndIf


		;-------------------------------------------------------------------------------
		; Setup HKCU registry settings for .VQX file association with VM.EXE (+1.4.1.2)
		;-------------------------------------------------------------------------------
		; Remove .vqx association in HKCU in favor of HKLM as setup by VEStartup
		RegDelete("HKEY_CURRENT_USER\.vqx")

		; Replace any "Open With" list
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vqx")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vqx\OpenWithList", "a", "REG_SZ", "vq.exe")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vqx\OpenWithList", "b", "REG_SZ", "NOTEPAD.EXE")
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vqx\OpenWithList", "MRUList", "REG_SZ", "ab")


		;------------------------------------------------
		; Setup Visual Quality Workflow Email (+1.4.1.4)
		;------------------------------------------------
		; Inherit settings from VM Email to VQ Email
		If String($VQEmailUseVMSettings) <> "{key_not_defined}" Then
			If $VQEmailUseVMSettings Then

				$RRTmp = RegRead($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "Auto Discover")
				If Not @error Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "Auto Discover", "REG_DWORD", $RRTmp)

				$RRTmp = RegRead($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP Server")
				If Not @error Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP Server", "REG_SZ", $RRTmp)

				$RRTmp = RegRead($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP Email Address")
				If Not @error Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP Email Address", "REG_SZ", $RRTmp)

				$RRTmp = RegRead($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "Exchange User")
				If Not @error Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "Exchange User", "REG_SZ", $RRTmp)

				$RRTmp = RegRead($Reg_HKCU_Base & "\VISUAL Manufacturing\Email", "SMTP UserName")
				If Not @error Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP UserName", "REG_SZ", $RRTmp)

			EndIf
		EndIf

		; Setup VQEmail settings
		If String($VQEmailAutoDiscover) <> "{key_not_defined}" Then
			If $VQEmailAutoDiscover Then
				RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "Auto Discover", "REG_DWORD", "0x00000001")
			Else
				RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "Auto Discover", "REG_DWORD", "0x00000000")
			EndIf
		EndIf

		If $VQEmailSMTPServer <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP Server", "REG_SZ", $VQEmailSMTPServer)
		If $VQEmailSMTPEmailAddress <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP Email Address", "REG_SZ", $VQEmailSMTPEmailAddress)
		If $VQEmailExchangeUser <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "Exchange User", "REG_SZ", $VQEmailExchangeUser)
		If $VQEmailDisplayName <> "{key_not_defined}" Then
			If $VQEmailDisplayName = "" Then
				RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP UserName", "REG_SZ", @UserName & "@" & @ComputerName)
			Else
				RegWrite($Reg_HKCU_Base & "\VISUAL Quality\Email", "SMTP UserName", "REG_SZ", $VQEmailDisplayName)
			EndIf
		EndIf


	EndIf ; VQ_Installed()

EndIf ; VMFG_Installed() Or VQ_Installed()

;---------------------------------------------------------------------------------------------
; Initialize CRM user configuration
;---------------------------------------------------------------------------------------------
If CRM_Installed() Then ;Visual CRM is installed

	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Configuring Visual CRM user registry settings...")
	Sleep($GUI_Delay)

	; Setup defaults for ServerNames and Databases
	If ($CRMDatabases[0] = 1 And $CRMDatabases[1] = "{key_not_defined}") Then $CRMDatabases[0] = 0
	If ($CRMServerNames[0] = 1 And $CRMServerNames[1] = "{key_not_defined}") Then $CRMServerNames[0] = 0
	If $CRMDatabases[0] <> 0 Then
		RegDelete($Reg_HKCU_Base & "\Visual CRM\Login\Databases")
		For $i = 1 To $CRMDatabases[0]
			If $CRMDBTypes[0] = $CRMDatabases[0] Then
				RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\Databases", StringStripWS($CRMDatabases[$i], 3), "REG_SZ", StringStripWS($CRMDBTypes[$i], 3))
			Else
				RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\Databases", StringStripWS($CRMDatabases[$i], 3), "REG_SZ", StringStripWS($CRMDBTypes[1], 3))
			EndIf
		Next
	EndIf
	If $CRMServerNames[0] <> 0 Then
		RegDelete($Reg_HKCU_Base & "\Visual CRM\Login\ServerNames")
		For $i = 1 To $CRMServerNames[0]
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\ServerNames", $CRMServerNames[$i], "REG_SZ", "")
		Next
	EndIf
	If $CRMOLEDBProviders[0] <> 0 Then
		RegDelete($Reg_HKCU_Base & "\Visual CRM\Login\OLEDBProviders")
		For $i = 1 To $CRMOLEDBProviders[0]
			$OLETmp = StringSplit($CRMOLEDBProviders[$i], "\")
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\OLEDBProviders", $OLETmp[1], "REG_SZ", $OLETmp[2])
		Next
	EndIf
	$RegTmp = RegRead($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "User")
	If $RegTmp = "" Then ;CRM has no history
		RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "DatabaseType", "REG_SZ", $CRMDBType)
		RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "OLEDBProvider", "REG_SZ", $CRMOLEDBProvider)
		RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "Details", "REG_DWORD", "0x00000000")
		If $CRMDatabase <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "Database", "REG_SZ", $CRMDatabase)
		If $CRMServer <> "{key_not_defined}" Then RegWrite($Reg_HKCU_Base & "\Visual CRM\Login\LastLogin", "Server", "REG_SZ", $CRMServer)
	EndIf

	;+1.8.0.x{
	; Added capture of existing Layout and Merge files for new profiles
	$RegTmp = CleanPath(StringLower(RegRead($Reg_HKCU_Base & "\Visual CRM\Options", "LayoutDir")))
	If $RegTmp <> "" Then ;LayoutDir present
		If $RegTmp <> StringLower($VEProfile_Path & "\CRM\Layout") Then ; LayourDir is not the VEProfiles path
			If FileExists($RegTmp) Then ; Existing LayoutDir is a valid location
				DirCopy($RegTmp, $VEProfile_Path & "\CRM\Layout", 1 + 8)
			EndIf
		EndIf
	EndIf
	$RegTmp = CleanPath(StringLower(RegRead($Reg_HKCU_Base & "\Visual CRM\Options", "Merge_DocumentDirectory")))
	If $RegTmp <> "" Then ;Merge_DocumentDirectory present
		If $RegTmp <> $VEProfile_Path & "\CRM\Merge Docs" Then ; Merge_DocumentDirectory is not the VEProfiles path
			If FileExists($RegTmp) Then ; Existing Merge_DocumentDirectory is a valid location
				DirCopy($RegTmp, $VEProfile_Path & "\CRM\Merge Docs", 1 + 8)
			EndIf
		EndIf
	EndIf
	$RegTmp = CleanPath(StringLower(RegRead($Reg_HKCU_Base & "\Visual CRM\Options", "Merge_TempDirectory")))
	If $RegTmp <> "" Then ;LayoutDir present
		If $RegTmp <> $VEProfile_Path & "\CRM\Merge Temp" Then ; Merge_TempDirectory is not the VEProfiles path
			If FileExists($RegTmp) Then ; Existing Merge_TempDirectory is a valid location
				DirCopy($RegTmp, $VEProfile_Path & "\CRM\Merge Temp", 1 + 8)
			EndIf
		EndIf
	EndIf
	;}+1.8.0.x

	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "LayoutDir", "REG_SZ", $VEProfile_Path & "\CRM\Layout")
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "Merge_DocumentDirectory", "REG_SZ", $VEProfile_Path & "\CRM\Merge Docs")
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "Merge_TempDirectory", "REG_SZ", $VEProfile_Path & "\CRM\Merge Temp")
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "VMPath", "REG_SZ", $VMInstallPath)
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "VQPath", "REG_SZ", $VQInstallPath)

	;+1.5.0.x
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "OnlineBooks", "REG_SZ", CleanPath(RegRead($Reg_HKLM_Base & "\Visual CRM", "InstallDir")) & "\help")
	If $CRMUseVMXforVMLaunch = "True" Then
		RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "UseVMX", "REG_DWORD", "0x00000001")
	ElseIf $CRMUseVMXforVMLaunch = "False" Then
		RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "UseVMX", "REG_DWORD", "0x00000000")
	EndIf
	RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "VMXPath", "REG_SZ", RegRead($Reg_HKCU_Base & "\Visual Manufacturing\Configuration", "Local Directory"))

	If $CRMUniqueLayouts <> "{key_not_defined}" Then
		If $CRMUniqueLayouts = "True" Then
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "UniqueLayouts", "REG_DWORD", "0x00000001")
		Else
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "UniqueLayouts", "REG_DWORD", "0x00000000")
		EndIf
	EndIf
	;+1.4.6.x
	If $CRMEmployOutlookForm <> "{key_not_defined}" Then
		If $CRMEmployOutlookForm = "True" Then
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "EmployOutlookForm", "REG_DWORD", "0xFFFFFFFF")
		Else
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "EmployOutlookForm", "REG_DWORD", "0x00000000")
		EndIf
	EndIf
	If $CRMHidePreviewPane <> "{key_not_defined}" Then
		If $CRMHidePreviewPane = "True" Then
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "HidePreviewPane", "REG_DWORD", "0xFFFFFFFF")
		Else
			RegWrite($Reg_HKCU_Base & "\Visual CRM\Options", "HidePreviewPane", "REG_DWORD", "0x00000000")
		EndIf
	EndIf

EndIf

;---------------------------------------------------------------------------------------------
; Update VE.Net environment if installed
;---------------------------------------------------------------------------------------------
; Test for presence of VE.Net
If $DotNetInstallPath <> "" Then

	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Configuring Visual .Net Applications user registry settings...")
	Sleep($GUI_Delay)

	; Create or Force VE.Net HKCU Registry key updates.
	$DefaultKey = RegRead($Infor_HKCU_Base & "\Data", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Data", "UseServer", "REG_SZ", "0")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Lsa", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		; 1.02:	Reports from server
		RegWrite($Infor_HKCU_Base & "\Lsa", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
		RegWrite($Infor_HKCU_Base & "\Lsa", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Lsa", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Lsa", "InstanceGroup", "REG_SZ", "_default")
		;----------------------------------------
		If $DotNetlsa_SaveSettingsOnExitFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Lsa", "SaveSettingsOnExit", "REG_SZ", "1")
		Else
			RegWrite($Infor_HKCU_Base & "\Lsa", "SaveSettingsOnExit", "REG_SZ", "0")
		EndIf
		;----------------------------------------
		If $DotNetlsa_ClearOnSaveFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Lsa", "ClearOnSave", "REG_SZ", "1")
		Else
			RegWrite($Infor_HKCU_Base & "\Lsa", "ClearOnSave", "REG_SZ", "0")
		EndIf
		;----------------------------------------
		If $DotNetlsa_InsertOnEndOfGridFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Lsa", "InsertOnEndOfGrid", "REG_SZ", "1")
		Else
			RegWrite($Infor_HKCU_Base & "\Lsa", "InsertOnEndOfGrid", "REG_SZ", "0")
		EndIf
		;----------------------------------------
		If $DotNetlsa_ShowErrorMsgBoxFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Lsa", "ShowErrorMsgBox", "REG_SZ", "1")
		Else
			RegWrite($Infor_HKCU_Base & "\Lsa", "ShowErrorMsgBox", "REG_SZ", "0")
		EndIf
		;----------------------------------------
		If $DotNetlsa_LogErrorsFlag = "True" Then
			RegWrite($Infor_HKCU_Base & "\Lsa", "LogErrors", "REG_SZ", "1")
		Else
			RegWrite($Infor_HKCU_Base & "\Lsa", "LogErrors", "REG_SZ", "0")
		EndIf
		;----------------------------------------
	EndIf

	If $DotNetAutoUpdateURL <> "key_not_found" Then
		; Always force AutoUpdateURL
		RegWrite($Infor_HKCU_Base & "\Lsa", "RemotePath", "REG_SZ", $DotNetAutoUpdateURL)
	EndIf

	If $DotNetReportsPathBase <> "key_not_found" Then
		; Always force ReportsPath
		RegWrite($Infor_HKCU_Base & "\Lsa", "ReportsPath", "REG_SZ", $DotNetReportsPathBase & "\Lsa\Reports")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\LsaDBUtility", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\LsaDBUtility", "InstanceGroup", "REG_SZ", "_default")
		RegWrite($Infor_HKCU_Base & "\LsaDBUtility", "WindowBounds", "REG_SZ", "120, 86, 800, 640")
		RegWrite($Infor_HKCU_Base & "\LsaDBUtility", "WindowState", "REG_SZ", "Normal")
		RegWrite($Infor_HKCU_Base & "\LsaDBUtility", "SplitPosition", "REG_SZ", "180")
		RegWrite($Infor_HKCU_Base & "\LsaDBUtility", "ContextTreeColumnWidth", "REG_SZ", "200")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Misc\Lsa\Vta\Kiosk\InstanceGroupToUseForm", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Misc\Lsa\Vta\Kiosk\InstanceGroupToUseForm")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Misc\Lsa\Vta\Kiosk\LoginDialog", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Misc\Lsa\Vta\Kiosk\LoginDialog")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vcrm", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vcrm", "ReportsPath", "REG_SZ", "$INSTALLPATH\Vcrm\Reports")
		RegWrite($Infor_HKCU_Base & "\Vcrm", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vcrm", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
		RegWrite($Infor_HKCU_Base & "\Vcrm", "LocalPath", "REG_SZ", "$INSTALLPATH")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vfin", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vfin", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vfin", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vfin", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
	EndIf
	If $DotNetReportsPathBase <> "key_not_found" Then
		; Always force ReportsPath
		RegWrite($Infor_HKCU_Base & "\Vfin", "ReportsPath", "REG_SZ", $DotNetReportsPathBase & "\Vfin\Reports")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vm", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vm", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
		RegWrite($Infor_HKCU_Base & "\Vm", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vm", "ReportsPath", "REG_SZ", "$INSTALLPATH\Vm\Reports")
		RegWrite($Infor_HKCU_Base & "\Vm", "BinPath", "REG_SZ", "$STARTUPPATH")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vmfg", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vmfg", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vmfg", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vmfg", "ReportsPath", "REG_SZ", "$INSTALLPATH\Vmfg\Reports")
		RegWrite($Infor_HKCU_Base & "\Vmfg", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vrm", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vrm", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vrm", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
		RegWrite($Infor_HKCU_Base & "\Vrm", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vrm", "ReportsPath", "REG_SZ", "$INSTALLPATH\Vrm\Reports")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vscp", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vscp", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vscp", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vscp", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
		RegWrite($Infor_HKCU_Base & "\Vscp", "ReportsPath", "REG_SZ", "$INSTALLPATH\Vscp\Reports")
	EndIf

	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vta", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vta", "LocalPath", "REG_SZ", "$INSTALLPATH")
		RegWrite($Infor_HKCU_Base & "\Vta", "BinPath", "REG_SZ", "$STARTUPPATH")
		RegWrite($Infor_HKCU_Base & "\Vta", "InstallPath", "REG_SZ", "$STARTUPPATH\..")
	EndIf


	If $DotNetReportsPathBase <> "key_not_found" Then
		; Always force ReportsPath
		RegWrite($Infor_HKCU_Base & "\Vta", "ReportsPath", "REG_SZ", $DotNetReportsPathBase & "\Vta\Reports")
	EndIf


	$DefaultKey = RegRead($Infor_HKCU_Base & "\Vta\Kiosk\EmployeeIdentificationForm", "")
	If @error = 1 Or $DotNetOverwriteFlag = "True" Then
		RegWrite($Infor_HKCU_Base & "\Vta\Kiosk\EmployeeIdentificationForm")
	EndIf

EndIf ;If $DotNetInstallPath <> "")


#cs ---------------------------------------------------------------------------------------------
	Include File:	VEProfilePolicy.au3
	Customer:       Synergy Resources, LLC

	AutoIt Version: 3.2.6.0
	Author:         Craig Gunst

	Included Function:
	MergeINIFiles ($SourceFile, $DestFile)
	Function to Merge two .ini files


	Copyright:	Synergy Resources, LLC Central Islip, NY. All rights reserved.
	Distribution of this intelectual property outside of the customers
	organization 	it was developed for, without the expressed permission
	of Synergy Resources, LLC is strictly prohibited.

	---------------------------------------------------------------------------------------------
	Releases:
	2007-10-11/cdg	Initial release.

#ce ---------------------------------------------------------------------------------------------

;-------------
; Definitions
;-------------
Dim $PolicyBasePath
Dim $UserBasePath



;-----------------------------------------------------
; Merge Profiles
;-----------------------------------------------------

$PolicyBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & "Policy Profile"
If $VEProfileDomain = "Local-PC-Logons" Then
	$UserBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & StringReplace(@UserName, ".", "_") & "_" & $VEProfileDomain
Else
	$UserBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & StringReplace(@UserName, ".", "_") & "_" & $VEProfileDomain
EndIf
$MachineBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\" & $VEProfileDomain & " Policy Profile"

; Update GUI message area
GUICtrlSetFont($guiMsgBox, -1, 400)
GUICtrlSetData($guiMsgBox, "Configuring Visual Enterprise user profile policy settings...")

If FileExists($UserBasePath) Then
	If FileExists($PolicyBasePath) Then
		ProfilePolicyMerge($PolicyBasePath, $UserBasePath)
		If FileExists($MachineBasePath) Then
			ProfilePolicyMerge($MachineBasePath, $UserBasePath)
		EndIf
		If $VEVELogonADGroupProfiles[1] <> "" Then

			GUICtrlSetFont($guiMsgBox, -1, 400)
			GUICtrlSetData($guiMsgBox, "Checking " & @UserName & " account for Active Directory group membership...")

			$ADGroupProfiles = _GetADUserGroups($VEVELogonADGroupProfiles)

			If $ADGroupProfiles[1] <> "" Then
				For $i = 1 To $ADGroupProfiles[0]
					$ADGroupBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain & "\ADGroup Profiles\{" & $ADGroupProfiles[$i] & "} Policy Profile"
					If FileExists($ADGroupBasePath) Then
						ProfilePolicyMerge($ADGroupBasePath, $UserBasePath)
					EndIf
				Next
			EndIf
		EndIf
	EndIf

	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Configuring Visual Enterprise user profile policy settings...")
	Sleep($GUI_Delay)

	; Check for VMBTS Log files and purge old files
	If $VEVELogonPurgeVMBTSLogDays <> 0 Then
		$VMBTSHndl = FileFindFirstFile($UserBasePath & "\VMFG\VMBTS*.LOG")
		If $VMBTSHndl = -1 Or @error Then
			;"No files/directories matched the search pattern")
		Else
			; Update GUI message area
			GUICtrlSetFont($guiMsgBox, -1, 400)
			GUICtrlSetData($guiMsgBox, "Purging old VMBTS log files from user profile...")
			Sleep($GUI_Delay)

			$SystemDate = @YEAR & "/" & @MON & "/" & @MDAY
			While 1
				; Get next search hit
				$VMBTSNextLog = FileFindNextFile($VMBTSHndl)
				If @error Then
					; No more hits
					ExitLoop
				Else
					;-------------------------
					; Process next search hit
					;-------------------------
					$FileDate = FileGetTime($UserBasePath & "\VMFG\" & $VMBTSNextLog, 0, 1)
					$FileDate = StringMid($FileDate, 1, 4) & "/" & StringMid($FileDate, 5, 2) & "/" & StringMid($FileDate, 7, 2)
					If Abs(_DateDiff("D", $SystemDate, $FileDate)) > $VEVELogonPurgeVMBTSLogDays Then
						FileDelete($UserBasePath & "\VMFG\" & $VMBTSNextLog)
					EndIf
				EndIf
			WEnd

		EndIf
		FileClose($VMBTSHndl)

	EndIf

EndIf

;-------------------------------------------------------
; Administrative profiles update:
;	Batch update all user profiles when /a parameter
;-------------------------------------------------------
If $CmdLine[0] > 0 Then
	If StringLower($CmdLine[1]) = "/a" Then
		$ProfilesBasePath = $VELogonVEProfilesPath & "\" & $VEProfileDomain
		$FLHndla = FileFindFirstFile($ProfilesBasePath & "\*.*")
		If $FLHndla = -1 Or @error Then
			;"No files/directories matched the search pattern")
		Else

			While 1
				; Get next search hit
				$FLNexta = FileFindNextFile($FLHndla)
				If @error Then
					; No more hits
					ExitLoop
				Else
					;-------------------------
					; Process next search hit
					;-------------------------
					If StringInStr(FileGetAttrib($ProfilesBasePath & "\" & $FLNexta), "D") <> 0 Then
						If StringLower($FLNexta) <> "policy profile" Then
							ProfilePolicyMerge($PolicyBasePath, $ProfilesBasePath & "\" & $FLNexta)
						EndIf
					EndIf
				EndIf
			WEnd
		EndIf
		FileClose($FLHndla)
	EndIf
EndIf



RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Normal")
Exit

;================================================================================================
; Function to recurse through $PolicyBasePath path and update/copy all file in/to $UserBasePath
; path. Except for .ini files, if the file in the destination has the same last-modified time as
; the same source file, the file is left alone, otherwise the file source is copied to the
; destination. .ini files in the source path are merged with the corisponding .ini file in the
; destination by the rules described in the MergeINIFiles function included.
;================================================================================================
Func ProfilePolicyMerge($PolicyBasePath, $UserBasePath)
	Local $FLHndl

	$FLHndl = FileFindFirstFile($PolicyBasePath & "\*.*")
	If $FLHndl = -1 Or @error Then
		;"No files/directories matched the search pattern")
		Return
	EndIf
	;-------------------------------------------------------
	; Recurse through directory and process any search hits
	;-------------------------------------------------------
	While 1
		; Get next search hit
		$FLNext = FileFindNextFile($FLHndl)
		If @error Then
			; No more hits
			ExitLoop
		Else
			;-------------------------
			; Process next search hit
			;-------------------------
			If StringInStr(FileGetAttrib($PolicyBasePath & "\" & $FLNext), "D") <> 0 Then
				;Search hit is a directory so recurse
				$DirExists = FileExists($UserBasePath & "\" & $FLNext)
				If $DirExists = 0 Then
					DirCreate($UserBasePath & "\" & $FLNext)
				EndIf
				ProfilePolicyMerge($PolicyBasePath & "\" & $FLNext, $UserBasePath & "\" & $FLNext)
			Else
				If StringLower(StringRight($FLNext, 4)) = ".ini" Then
					GUICtrlSetData($guiMsgBox, "Configuring Visual Enterprise user profile policy settings... merging " & $FLNext)
					MergeINIFiles($PolicyBasePath & "\" & $FLNext, $UserBasePath & "\" & $FLNext)
				Else
					GUICtrlSetData($guiMsgBox, "Configuring Visual Enterprise user profile policy settings... copying " & $FLNext)
					MergeProfile($PolicyBasePath & "\" & $FLNext, $UserBasePath & "\" & $FLNext)
				EndIf
			EndIf
		EndIf
	WEnd

	FileClose($FLHndl)

EndFunc   ;==>ProfilePolicyMerge

Func MergeProfile($PolicyProfileItem, $UserProfileItem)

	$PPITime = FileGetTime($PolicyProfileItem, 0, 1)
	$UPITime = FileGetTime($UserProfileItem, 0, 1)

	; Update/copy non .ini files in user profile
	If $VEVELogonProfileReadOnlyCopyFlag > 0 Then
		$aPolicyProfileItem = FileParse($PolicyProfileItem)
		For $aPos = 1 To $VEVELogonProfileReadOnlyCopyExt[0]
			If $aPolicyProfileItem[4] = $VEVELogonProfileReadOnlyCopyExt[$aPos] Then
				If $PPITime <> $UPITime Then
					Select
						Case $VEVELogonProfileReadOnlyCopyFlag = 1
							FileCopyRO($PolicyProfileItem, $UserProfileItem, 1 + 8)
							ExitLoop
						Case $VEVELogonProfileReadOnlyCopyFlag = 2
							FileCopyRO($PolicyProfileItem, $UserProfileItem, 1 + 8 + 16)
							ExitLoop
						Case $VEVELogonProfileReadOnlyCopyFlag = 3
							FileCopyRO($PolicyProfileItem, $UserProfileItem, 1 + 8 + 32)
							ExitLoop
					EndSelect
				Else
					If $VEVELogonProfileReadOnlyCopyFlag = 3 Then
						If StringInStr(FileGetAttrib($UserProfileItem), "D") = 1 Then
							If StringInStr(FileGetAttrib($UserProfileItem & "\" & $aPolicyProfileItem[2]), "R") = 0 Then
								FileSetAttrib($UserProfileItem & "\" & $aPolicyProfileItem[2], "+R")
							EndIf
						Else
							If StringInStr(FileGetAttrib($UserProfileItem), "R") = 0 Then
								FileSetAttrib($UserProfileItem, "+R")
							EndIf
						EndIf
					EndIf
				EndIf

			Else
				If $aPos = $VEVELogonProfileReadOnlyCopyExt[0] Then
					FileCopy($PolicyProfileItem, $UserProfileItem, 1 + 8)
					ExitLoop
				EndIf
			EndIf
		Next
	Else
		FileCopy($PolicyProfileItem, $UserProfileItem, 1 + 8)
	EndIf

EndFunc   ;==>MergeProfile


;=========================================================================================================================
; Function to Merge two .ini files
;=========================================================================================================================
;	SYNTAX:
;	---------------------------------------
;	MergeINIFiles ($SourceFile, $DestFile)
;
;
;	DESCRIPTION:
;	--------------------------------------------------------------------------------------------
;	->	Each keys=value in the source .ini file will be copied to the destination .ini file.
;	->	If the key=value exists in the destination .ini it will be overwritten by the key=value
;		in the source.
;	->	A caret (^) at the end of the section name in the source .ini will cause the entire section
;		(section name less the caret) to be deleted prior to the copying of any key=value's.
;	->	To delete the section completely you must include one key=value with the key delete
;		syntax, i.e. "Section_Deleted^=".
;	->	A caret (^) at the end of a key name will delete that key=value from the destination .ini.
;		(same as "key=" without a value except the key is also removed)
;
;
;	EXAMPLE: (ellipsis "..." represent deleted or missing lines for readability)
;	--------------------------------------------------------------------------------------------------------------------
;	SourceFile						Destination File Before						Destination File After
;	----------------------------	----------------------------------------	----------------------------------------
;	[Visual Mfg]					[Visual Mfg]								[Visual Mfg]
;	Database=SRI					Database=VTG								Database=SRI
;	UserID=							UserID=JoeUser								UserID=
;	Password^=						Password=secret								...
;
;	[Semaphore^]					[Semaphore]									[Semaphore]
;	Engineering=0					Engineering=1								Engineering=0
;	WorkOrder=0						WorkOrder=1									WorkOrder=0
;	Inventory=0						Inventory=1									Inventory=0
;	...								Labor=1										...
;	...								Mrp=1										...
;	...								Receivables=1								...
;
;	[CheckGeneration^]				[CheckGeneration]							...
;	Secttion_Deleted^=				CheckStyle=T								...
;	...								NumberOfStubs=2								...
;
;	[GLReportWriter]				...											[GLReportWriter]
;	SavePath=\\Server\Share\Path	...											SavePath=\\Server\Share\Path
;
;	[Shipping]						[Shipping]									[Shipping]
;	...								NotationPosLine=774,38,1024,212				NotationPosLine=774,38,1024,212
;	...								NotationPosCustOrderLine=732,212,1024,386	NotationPosCustOrderLine=732,212,1024,386
;	...								NotationPosHeader=766,280,1016,454			NotationPosHeader=766,280,1016,454
;	SuppressZeroLines=Yes			...											SuppressZeroLines=Yes
;
;=========================================================================================================================

Func MergeINIFiles($SourceFile, $DestFile)

	$IniSectNames = IniReadSectionNames($SourceFile)
	If @error Then Return False
	For $NextSect = 1 To $IniSectNames[0]

		$IniKeys = IniReadSection($SourceFile, $IniSectNames[$NextSect])
		; + 1.4.0.2 {
		; If section exists but is empty skip to the next.
		If @error = 1 Then ContinueLoop
		; }
		If StringRight($IniSectNames[$NextSect], 1) = "^" Then
			$IniSectNames[$NextSect] = StringTrimRight($IniSectNames[$NextSect], 1)
			IniDelete($DestFile, $IniSectNames[$NextSect])
		EndIf

		For $NextKey = 1 To $IniKeys[0][0]
			If StringRight($IniKeys[$NextKey][0], 1) = "^" Then
				IniDelete($DestFile, $IniSectNames[$NextSect], StringTrimRight($IniKeys[$NextKey][0], 1))
			Else
				IniWrite($DestFile, $IniSectNames[$NextSect], $IniKeys[$NextKey][0], $IniKeys[$NextKey][1])
			EndIf
		Next

	Next
	Return True

EndFunc   ;==>MergeINIFiles


Func VMFG_Installed()
	; {+1.8.1.x
	; If VMInstalledClue is defined look for valid path else use HKLM technique
	If $VELogonVMInstalledClue[1] <> "" And $VELogonVMInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VELogonVMInstalledClue[0]
			If FileExists($VELogonVMInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf (StringInStr($RegKey, "Visual") > 0 And (StringInStr($RegKey, "Manufacturing") > 0 Or StringInStr($RegKey, "Mfg") > 0 Or StringInStr($RegKey, "Enterprise") > 0)) Or (StringInStr($RegKey, "Infor") > 0 And (StringInStr($RegKey, "10") > 0 And StringInStr($RegKey, "ERP") > 0 And StringInStr($RegKey, "Express") > 0)) Then
			Return True
		EndIf
	WEnd
	If $VEVELogonPre650VisualFlag = "True" Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual") > 0 And (StringInStr($RegKey, "Manufacturing") > 0 Or StringInStr($RegKey, "Mfg") > 0 Or StringInStr($RegKey, "Enterprise") > 0) Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>VMFG_Installed

Func CRM_Installed()
	; {+1.8.1.x
	; If CRMInstalledClue is defined look for valid path else use HKLM technique
	If $VELogonCRMInstalledClue[1] <> "" And $VELogonCRMInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VELogonCRMInstalledClue[0]
			If FileExists($VELogonCRMInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf StringInStr($RegKey, "Visual CRM") Then
			Return True
		EndIf
	WEnd
	If $VEVELogonPre650VisualFlag = "True" Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual CRM") Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>CRM_Installed

Func VQ_Installed()
	; {+1.8.1.x
	; If VQInstalledClue is defined look for valid path else use HKLM technique
	If $VELogonVQInstalledClue[1] <> "" And $VELogonVQInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VELogonVQInstalledClue[0]
			If FileExists($VELogonVQInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf StringInStr($RegKey, "Visual Quality") Then
			Return True
		EndIf
	WEnd
	If $VEVELogonPre650VisualFlag = "True" Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual Quality") Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>VQ_Installed






;--------------------------------------------------------------------------------------
; GetIniArray - Get multi-value key from .ini file
;--------------------------------------------------------------------------------------
; Function to read multiple entries from .ini key to array
; array[0] contains the number of elements, array[1-array[0]]
; contains the values stripped of leading and trailing space

Func GetIniArray($iniFilename, $iniSection, $iniKey, $iniDefault)

	Local $IniReturn[32], $IniTemp, $i

	$IniTemp = IniRead($iniFilename, $iniSection, $iniKey, "~{KeyNotFound}")
	If $IniTemp = "~{KeyNotFound}" Then
		$IniTemp = $iniDefault
	EndIf

	$IniReturn = StringSplit($IniTemp, ",", 0)
	For $i = 1 To $IniReturn[0]
		$IniReturn[$i] = StringStripWS($IniReturn[$i], 3)
	Next

	Return $IniReturn

EndFunc   ;==>GetIniArray

; Added in 1.3.4.0	/cdg {
;--------------------------------------------------------------------------------------
; CleanExt - Remove leading period (".") from file extension list
;--------------------------------------------------------------------------------------
; Also strips leading and trailing space
; If variable is an array, array[0]= number of elements

Func CleanExt($InputExt)

	Local $i
	If IsArray($InputExt) Then
		For $i = 1 To $InputExt[0]
			$InputExt[$i] = StringStripWS($InputExt[$i], 3)
			If StringLeft($InputExt[$i], 1) = "." Then
				$InputExt[$i] = StringRight($InputExt[$i], StringLen($InputExt[$i]) - 1)
			EndIf
		Next
	Else
		$InputExt = StringStripWS($InputExt, 3)
		If StringLeft($InputExt, 1) = "." Then
			$InputExt = StringRight($InputExt, StringLen($InputExt) - 1)
		EndIf
	EndIf

	Return $InputExt

EndFunc   ;==>CleanExt
; }

;--------------------------------------------------------------------------------------
; CleanPath - Remove trailing backslash ("\") from path
;--------------------------------------------------------------------------------------
; Also strips leading and trailing space
; If variable is an array, array[0]= number of elements

Func CleanPath($InputPath)

	Local $i
	If IsArray($InputPath) Then
		For $i = 1 To $InputPath[0]
			$InputPath[$i] = StringStripWS($InputPath[$i], 3)
			If StringRight($InputPath[$i], 1) = "\" Then
				$InputPath[$i] = StringLeft($InputPath[$i], StringLen($InputPath[$i]) - 1)
			EndIf
		Next
	Else
		$InputPath = StringStripWS($InputPath, 3)
		If StringRight($InputPath, 1) = "\" Then
			$InputPath = StringLeft($InputPath, StringLen($InputPath) - 1)
		EndIf
	EndIf

	Return $InputPath

EndFunc   ;==>CleanPath


;--------------------------------------------------------------------------------------
; GetIniFlag - Read ini flag parameter Yes/No and return True/False
;--------------------------------------------------------------------------------------

Func GetIniFlag($GIFIniFile, $GIFIniSection, $GIFIniKey, $GIFDefault)

	$GIFTemp = IniRead($GIFIniFile, $GIFIniSection, $GIFIniKey, "{key_not_defined}")
	If $GIFTemp = "{key_not_defined}" Then
		Return $GIFDefault
	EndIf

	$GIFTemp = StringStripWS(StringLower($GIFTemp), 3)
	If StringInStr("|1|yes|y|true|t|", $GIFTemp) > 0 Then Return "True"
	If StringInStr("|0|no|n|false|f|", $GIFTemp) > 0 Then Return "False"
	MsgBox(0, "VELogon - Error: Value Out of Range", _
			$ScriptFile_ini & @CRLF & _
			"""" & $GIFIniKey & """ key in the ""[" & $GIFIniSection & "]"" section of the" & @CRLF & _
			"""" & $GIFIniFile & """ file has an invalid value specified" & @CRLF & _
			"Valid positive values are: 1, y, yes, t or true" & @CRLF & _
			"Valid negative values are: 0, n, no, f or false" & @CRLF & _
			"Values are not case sensative")
	RegWrite("HKEY_CURRENT_USER\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", $GIFIniSection & " / " & $GIFIniKey & " Value Out of Range")
	Exit

EndFunc   ;==>GetIniFlag


;--------------------------------------------------------------------------------------
; _GetADUserGroups - Read Active Directory profile security group membership
;--------------------------------------------------------------------------------------

Func _GetADUserGroups($ADGroups)

	Local $objUser, $CurrentUser
	Local $vMembershipList = ""

	$objUser = ObjCreate("ADSystemInfo")
	$oCOMError = ObjEvent("AutoIt.Error", "_ADError")
	$CurrentUser = ObjGet("LDAP://" & $objUser.UserName)
	$avMembersList = $CurrentUser.MemberOf
	If IsArray($avMembersList) Then
		For $vMember In $avMembersList
			$avMemberParsed = StringSplit($vMember, ",")
			$avMemberParsed = StringSplit($avMemberParsed[1], "=")
			If $vMembershipList <> "" Then $vMembershipList &= ","
			$vMembershipList &= $avMemberParsed[2]
		Next
	Else
		$avMemberParsed = StringSplit($avMembersList, ",")
		$avMemberParsed = StringSplit($avMemberParsed[1], "=")
		If $avMemberParsed[0] > 1 Then
			$vMembershipList &= $avMemberParsed[2]
		Else
			$vMembershipList = ""
		EndIf
	EndIf

	$avMembershipList = _StringSplitTrim($vMembershipList, ",")

	Local $vADGList = ""

	If Not IsArray($ADGroups) Then
		If $ADGroups = "*" Then
			Return $avMembershipList
		Else
			$ADGroups = _StringSplitTrim($ADGroups, ",")
		EndIf
	EndIf

	For $iADG = 1 To $ADGroups[0]
		For $iML = 1 To $avMembershipList[0]
			If StringLower($ADGroups[$iADG]) = StringLower($avMembershipList[$iML]) Then
				If $vADGList <> "" Then $vADGList &= ","
				$vADGList &= $ADGroups[$iADG]
			EndIf
		Next
	Next
	$avADGList = _StringSplitTrim($vADGList, ",")
	Return $avADGList

EndFunc   ;==>_GetADUserGroups

Func _ADError()
	$HexNumber = Hex($oCOMError.number, 8)
	MsgBox(16, @ScriptName & " - Error", "Unable to read Active Directory groups" & @CRLF & _
			"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
			"Please report this error to your System Administrator")
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to read Active Directory groups")
	Exit
EndFunc   ;==>_ADError


;--------------------------------------------------------------------------------------
; _StringSplitTrim - Split string and trim leading and trailing whitespace
;--------------------------------------------------------------------------------------

Func _StringSplitTrim($SST_String, $SST_Delimiter)

	Local $iSST
	$sSST_Tmp = StringSplit($SST_String, $SST_Delimiter)
	For $iSST = 1 To $sSST_Tmp[0]
		$sSST_Tmp[$iSST] = StringStripWS($sSST_Tmp[$iSST], 3)
	Next

	Return $sSST_Tmp
EndFunc   ;==>_StringSplitTrim

;============================================================================================
; Display GUI
;============================================================================================

Func GUIBuildMain()

	$iMainWidth = 500
	$iCenterHight = 0

	$iMainHight = $iCenterHight + 123 ; Minimum = 123 for no center window

	$iTitleHight = 100
	$iLogoWidth = 80

	;--------------------------------------------------------------------------------------------
	; Build Main Window
	;--------------------------------------------------------------------------------------------
	;	$GUID_Main = GUICreate($ScriptName & " - " & $vScriptDesc , $iMainWidth, $iMainHight, -1, -1, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	;	$GUID_Main = GUICreate($ScriptName & " - " & $vScriptDesc , $iMainWidth, $iMainHight, -1, -1, BitOR($WS_POPUPWINDOW,$WS_CAPTION),BitOR($WS_EX_TOPMOST  ,0) )
	$guiMain = GUICreate($ScriptName & " - " & $vScriptDesc, $iMainWidth, $iMainHight + 27, -1, -1, $WS_DLGFRAME)


	;--------------------------------------------------------------------------------------------
	; Build top banner
	;--------------------------------------------------------------------------------------------
	; Install Synergy Banner
	$guiLogo = GUICtrlCreatePic($ScriptTempDir & "\SynergyPH3.JPG", 0, 0, $iMainWidth, $iTitleHight)
	GUICtrlSetState(-1, $GUI_DISABLE)


	$V_Offset = 2
	$H_Offset = 5
	$Banner_Font = "Arial Black"
	GUICtrlCreateLabel($Banner_Title, $iLogoWidth + $H_Offset, 15 + $V_Offset, $iMainWidth - $iLogoWidth - $H_Offset, 70 - $V_Offset, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 400, 0, $Banner_Font)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetState(-1, $GUI_ENABLE)

	; Title Text
	GUICtrlCreateLabel($Banner_Title, $iLogoWidth, 15, $iMainWidth - $iLogoWidth, 70, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 400, 0, $Banner_Font)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0xFFFFFF)

	; Display copyright in lower-right of banner
	GUICtrlCreateLabel("©Synergy Resources, LLC", $iMainWidth - 150, 84, 140, 12, $SS_RIGHT)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0xB0B0B0)

	; Display center window
	If $iCenterHight > 0 Then
		GUICtrlCreateGraphic(2, $iTitleHight + 2, $iMainWidth - 4, ($iMainHight - $iTitleHight) - 24, $SS_SUNKEN)
		GUICtrlSetState(-1, $GUI_DISABLE)
	EndIf


	; Display program version in lower-left of window
	$ScriptRev = FileGetVersion(@ScriptFullPath)
	$iRevWidth = (StringLen($ScriptRev) * 5.5) + 24
	GUICtrlCreateGraphic(($iMainWidth - $iRevWidth) - 6, $iMainHight - 21, $iRevWidth + 5, 20, $SS_SUNKEN)
	GUICtrlCreateLabel("rev: " & $ScriptRev, ($iMainWidth - $iRevWidth) - 4, $iMainHight - 18, $iRevWidth, 14, $SS_CENTER)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetColor(-1, 0x505050)

	; Display Message Area in lower-left of window
	GUICtrlCreateGraphic(2, $iMainHight - 21, ($iMainWidth - $iRevWidth) - 9, 20, $SS_SUNKEN)
	$guiMsgBox = GUICtrlCreateLabel("", 6, $iMainHight - 18, ($iMainWidth - $iRevWidth) - 14, 14, $SS_LEFT)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetColor(-1, 0x505050)

EndFunc   ;==>GUIBuildMain

;--------------------------------------------------------------------------------------
; FileCopy of read-only file, i.e. overwrite read-only destination file (+1.4.4.x)
;
;	flag:	0	=	do not overwrite existing file
;			1	=	overwrite existing file
;			8	=	create destination directory structure if it doesn't exist
;			16	=	preserve state of destination read-only attribute
;			32	=	force destination to read-only attribute
;
;	if source is read-only destination will be read-only
;	if source is read-write destination will be read-write unless flag bit 16
;	is set in which case the state of the destination read-only attribute will
;	be preserved after the copy.
;
;--------------------------------------------------------------------------------------

Func FileCopyRO($FCROsource, $FCROdest, $FCROflag)

	Dim $FCROset = False
	If FileExists($FCROdest) = 1 Then
		; if destination is a directory check the file within
		If StringInStr(FileGetAttrib($FCROdest), "D") <> 0 Then
			$aFCROSource = FileParse($FCROsource)
			$FCROdest = CleanPath($FCROdest) & "\" & $aFCROSource[2]
		EndIf
		If FileExists($FCROdest) Then
			If BitAND($FCROflag, 1) = 1 Then
				If StringInStr(FileGetAttrib($FCROdest), "R") <> 0 Then
					$FCROset = True
					FileSetAttrib($FCROdest, "-R")
				EndIf
			EndIf
		EndIf
	EndIf
	$FCROresult = FileCopy($FCROsource, $FCROdest, BitAND($FCROflag, 1 + 8))
	If $FCROresult = 1 Then
		If BitAND($FCROflag, 32) = 32 Then
			FileSetAttrib($FCROdest, "+R")
		Else
			If BitAND($FCROflag, 16) = 16 Then
				If $FCROset = True Then
					FileSetAttrib($FCROdest, "+R")
				Else
					FileSetAttrib($FCROdest, "-R")
				EndIf
			EndIf
		EndIf
	EndIf

EndFunc   ;==>FileCopyRO


;--------------------------------------------------------------------------------------
; Parse path\filename (+1.4.4.x)
;--------------------------------------------------------------------------------------
; Return an array of the parsing of the filename
;	i.e. FileParse("c:\Visual\RunTime40\sql.ini") would return:
;	FileParse[0] = Drive ("c:")
;	FileParse[1] = Path ("\Visual\RunTime40\")
;	FileParse[2] = Filename ("sql.ini")
;	FileParse[3] = Filename First Name ("sql")
;	FileParse[4] = Filename Extension ("ini")

; 1 = path, 2 = filename, 3 = filename, first name, and 4 = filename extension
;--------------------------------------------------------------------------------------

Func FileParse($FPfile)
	Dim $FPReturn[5]
	$FPPos = StringInStr($FPfile, ":", 0, 1)
	If $FPPos = 0 Then
		$FPDrive = ""
	Else
		$FPDrive = StringLeft($FPfile, $FPPos)
		$FPfile = StringRight($FPfile, StringLen($FPfile) - $FPPos)
	EndIf

	$FPPos = StringInStr($FPfile, "\", 0, -1)
	If $FPPos = 0 Then
		$FPFilename = $FPfile
		$FPPath = ""
	Else
		$FPFilename = StringRight($FPfile, StringLen($FPfile) - $FPPos)
		$FPPath = StringLeft($FPfile, $FPPos)
	EndIf

	$FPPos = StringInStr($FPFilename, ".", 0, -1)
	If $FPPos = 0 Then
		$FPFileFirstname = $FPFilename
		$FPFileExtension = ""
	Else
		$FPFileFirstname = StringLeft($FPFilename, $FPPos - 1)
		$FPFileExtension = StringRight($FPFilename, StringLen($FPFilename) - $FPPos)
	EndIf

	$FPReturn[0] = $FPDrive
	$FPReturn[1] = $FPPath
	$FPReturn[2] = $FPFilename
	$FPReturn[3] = $FPFileFirstname
	$FPReturn[4] = $FPFileExtension

	Return $FPReturn

EndFunc   ;==>FileParse



