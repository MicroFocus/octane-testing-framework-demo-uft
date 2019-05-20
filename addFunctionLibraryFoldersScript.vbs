'************************************************************************************************************************
'Description:
'
'This example opens UFT and uses the Folders collection to configure
'the search paths that are used to resolve relative paths.
'
'Assumptions:
'There is no unsaved test currently open in UFT.
'https://admhelp.microfocus.com/uft/en/all/AutomationObjectModel/Default.htm#QuickTest~FoldersOptions~Add~Add%20a%20Folder%20to%20a%20Folder%20Search%20List_E.html?Highlight=folder
'************************************************************************************************************************

Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
Dim strPath
Dim scriptdir

' Open UFT
Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
qtApp.Launch ' Start UFT
qtApp.Visible = True ' Make the UFT application visible


' Locate "Folder1" and if it's not in the collection - add it
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strPath = qtApp.Folders.Locate(scriptdir + "\UftTestDemo\FeatureWithFunctionLibraries\functionLibraries")


' If returned empty string, then we cannot locate the absolute path, so there is nothing to do
If strPath <> "" Then
  If qtApp.Folders.Find(strPath) = -1 Then ' If the folder is not found in the collection
    qtApp.Folders.Add strPath, 1 ' Add the folder to the collection
  End If
End If

Set qtApp = Nothing ' Release the Application object