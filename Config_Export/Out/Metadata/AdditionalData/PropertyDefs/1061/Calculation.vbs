Option Explicit

Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal oFSO.OpenTextFile("C:\Temp\Scripts\MFUtils.vbs").ReadAll()
ExecuteGlobal oFSO.OpenTextFile("C:\Temp\Scripts\MFSign.vbs").ReadAll()
Set oFSO = Nothing

Dim oLog: Set oLog = New Logger
oLog.Path = "C:\temp\scripts\log.txt" 
Call oLog.Write(LOG_DEBUG, "Auto Calculate Signature " + ObjVerToStr(ObjVer) )


