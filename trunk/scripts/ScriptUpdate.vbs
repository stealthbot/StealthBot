#include "\ScriptUpdate\ut_Updater.vbs"
#include "\ScriptUpdate\frmScriptUpdate.vbs"
#include "\ScriptUpdate\frmScriptDownloadWizard.vbs"
#include "\lib\XMLTextWriter\XMLTextWriter.vbs"
#include "\lib\RssReader\RssReader.vbs"
#include "\lib\libFiftyToo\Functions.vbs"
#include "\lib\libFiftyToo\Color.vbs"



'// Change Log
'// 0.1 - Initial Draft
Option Explicit

Private Updater, mnuScriptUpdate

' // script data
Script("Name") = "ScriptUpdate"		' script name
Script("Author") = "FiftyToo"		' script author
Script("Major") = 0           		' script major version
Script("Minor") = 1           		' script minor version
Script("Revision") = 0        		' script version revision   

Sub Event_Load()
	
	'// Create our Updater instance
	Set Updater = New ut_Updater
	
	'// Create menus
	Call CreateObj("Menu", "mnuScriptUpdate")
	mnuScriptUpdate.Caption = "Download Scripts"
	Set frmScriptUpdate = Nothing
	
	'// for testing...
	'Call Updater.FetchSBScriptsInCategory(2)
	'Call Updater.FetchCategories()
	'Call Updater.CheckForUpdates()
	'Call Updater.SearchScripts(0, "update")
	
	'mnuScriptUpdate_Click()
	
End Sub

Sub mnuScriptUpdate_Click()

	'// Create forms
	If (frmScriptUpdate Is Nothing) = True Then
		Call CreateObj("Form", "frmScriptUpdate")
		Call CreateObj("Form", "frmScriptDownloadWizard")
	End If
	
	Call frmScriptUpdate.Show()
	
End Sub

Sub Event_PressedEnter(Text)

End Sub


Sub Event_UserTalk(Username, Flags, Message, Ping)

End Sub

     
