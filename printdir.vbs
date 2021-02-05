
Option Explicit
On Error Resume Next
 
Const WAIT_TIME  = 10000 '5 seconds
Const PRINT_TIME = 10000 '5 seconds
 
Dim WshShell, fso, configFile, objReadFile, str64, strPath, ApplicationData
Dim dbWatchDir, attFolder, objShell, objFolder, colItems, objItem, dbLogDir, logFolder, doneFolder
 
Set WshShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

dbWatchDir = "S:\" 'inserire il percorso della cartella da monitorare
 
If Not fso.FolderExists (dbWatchDir) Then
 Set attFolder = fso.CreateFolder (dbWatchDir)
 WScript.Echo "Created a watch folder to hold your incoming print jobs at " & dbWatchDir
End If
 
dbLogDir = dbWatchDir & "\PrintLog"
 
If Not fso.FolderExists (dbLogDir) Then
 Set logFolder = fso.CreateFolder (dbLogDir)
 'WScript.Echo "Created a folder to archive processed jobs - " & dbLogDir
End If
 
Do While True
 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(dbWatchDir)
Set colItems = objFolder.Items
doneFolder = dbLogDir & "\" 
 'doneFolder = dbLogDir & "\" & DateDiff("s", "1/1/2010", Now)
 
For Each objItem in colItems
 If Not objItem.IsFolder Then  
  If Not fso.FolderExists (doneFolder) Then
   Set logFolder = fso.CreateFolder (doneFolder)
   'WScript.Echo "Created a folder to archive processed jobs - " & doneFolder
  End If
  objItem.InvokeVerbEx("Print")
  'WScript.Echo "Now printing: " & objItem.Name  
  WScript.Sleep(PRINT_TIME)
  fso.MoveFile dbWatchDir & "\" & objItem.Name & "*", doneFolder
 end if
Next
WScript.Sleep(WAIT_TIME)
Set objShell = nothing
Set objFolder = nothing
Set colItems = nothing
Loop
 


