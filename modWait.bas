Attribute VB_Name = "modWait"
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function Wait(ByVal TimeToWait As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000
    Do Until GetTickCount > EndTime
        DoEvents
        Loop
End Function
