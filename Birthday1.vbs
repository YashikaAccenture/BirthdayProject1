 ' Create a WshShell to get the current directory
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Create an Excel instance
Dim myExcelWorker
Set myExcelWorker = CreateObject("Excel.Application") 

' Disable Excel UI elements
myExcelWorker.DisplayAlerts = False
myExcelWorker.AskToUpdateLinks = False
myExcelWorker.AlertBeforeOverwriting = False
myExcelWorker.FeatureInstall = msoFeatureInstallNone

' Tell Excel what the current working directory is 
' (otherwise it can't find the files)
Dim strSaveDefaultPath
Dim strPath
strSaveDefaultPath = myExcelWorker.DefaultFilePath
strPath = WshShell.CurrentDirectory
myExcelWorker.DefaultFilePath = strPath

' Open the Workbook specified on the command-line 
Dim oWorkBook
Dim strWorkerWB
strWorkerWB = strPath & "\YourWorkbook.xls"

Set oWorkBook = myExcelWorker.Workbooks.Open(strWorkerWB)

' Build the macro name with the full path to the workbook
Dim strMacroName
strMacroName = "'" & strPath & "\YourWorkbook" & "!Sheet1.YourMacro"
on error resume next 
   ' Run the calculation macro
   myExcelWorker.Run strMacroName
   if err.number <> 0 Then
      ' Error occurred - just close it down.
   End If
   err.clear
on error goto 0 

oWorkBook.Save 

myExcelWorker.DefaultFilePath = strSaveDefaultPath

' Clean up and shut down
Set oWorkBook = Nothing

' Don’t Quit() Excel if there are other Excel instances 
' running, Quit() will shut those down also
if myExcelWorker.Workbooks.Count = 0 Then
   myExcelWorker.Quit
End If

Set myExcelWorker = Nothing
Set WshShell = Nothing