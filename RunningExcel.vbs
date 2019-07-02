Option Explicit

Dim ObjExcel 
Dim ObjWB
On Error Resume Next


Set ObjExcel = GetObject(, "excel.application") 'gives error 429 if Word is not open
If Err.Number = 429 Then
  Err.Clear
  Set ObjExcel = CreateObject("excel.Application")
   
End If
If Not ObjExcel Is Nothing Then
   ObjExcel.Visible = True
Set ObjWB =ObjExcel.Workbooks.Open("C:\Users\asha.chauhan\OneDrive - Accenture\Birthday\Sheet\Birthday.xlsm")

Else
   Msgbox "Unable to retrieve Excel."
End If

ObjWB.Close False 
ObjExcel.Quit
Set ObjExcel = Nothing

Function PrintLog(argument)

Dim objFSO
dim objFile
dim thisLine
Dim FilePath
Dim log
FilePath = "C:\Users\yashika.a.gupta\Desktop\Birthday\Birthday\BirthdayTool" & ".txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")

If (objFSO.FileExists(FilePath)) Then
  Set log = objFSO.OpenTextFile(FilePath, ForWriting , True) 
  MsgBox "hell6"
Else
  Set log = objFSO.CreateTextFile(FilePath, TRUE)
End If
log.writeLine argument

End Function
