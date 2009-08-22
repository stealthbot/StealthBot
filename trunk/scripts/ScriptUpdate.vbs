#include "\ScriptUpdate\ut_Updater.vbs"
#include "\ScriptUpdate\frmScriptUpdate.vbs"
#include "\ScriptUpdate\frmScriptDownloadWizard.vbs"
#include "\lib\Stealthbot\XMLTextWriter.vbs"



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

	Call CreateScriptCommands()
	
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
	Call ShowUpdateForm()
End Sub

Sub Event_Command(cmd)
	Select Case LCase(cmd.Name)
		Case "updates":
			Call ShowUpdateForm()
	End Select
End Sub


Sub ShowUpdateForm()
	'// Create forms
	If (frmScriptUpdate Is Nothing) = True Then
		Call CreateObj("Form", "frmScriptUpdate")
		Call CreateObj("Form", "frmScriptDownloadWizard")
	End If
	
	Call frmScriptUpdate.Show()
End Sub


Sub CreateScriptCommands()

	Dim Command

    '// UPDATES
    Set Command = OpenCommand("updates")
    If Command Is Nothing Then
        '// It does not, lets create the command
        Set Command = CreateCommand("updates")
        With Command
			'// set a command description
			.Description = "Shows the script update form."
            '// save the command
            .Save
        End With
    End If


End Sub


     
