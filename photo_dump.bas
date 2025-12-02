Attribute VB_Name = "Module1"
Option Explicit

' Selected Emails Only ? saves images to Pictures\<sanitized subject>
Public Sub PictureDump_EmailsOnly()
    On Error GoTo EH

    Dim exp As Explorer
    Dim sel As Selection
    Dim itm As Object
    Dim firstMail As MailItem
    Dim target As String

    Set exp = Application.ActiveExplorer
    If exp Is Nothing Then
        MsgBox "No active Outlook window.", vbExclamation
        Exit Sub
    End If

    Set sel = exp.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Select one or more messages first.", vbInformation
        Exit Sub
    End If

    ' find the first MailItem in the selection
    Dim i As Long
    For i = 1 To sel.Count
        If TypeOf sel.Item(i) Is MailItem Then
            Set firstMail = sel.Item(i)
            Exit For
        End If
    Next i
    If firstMail Is Nothing Then
        MsgBox "No mail items in selection.", vbInformation
        Exit Sub
    End If

    target = BuildTargetFolderFromSubject(firstMail.subject)
    EnsureFolder target

    For Each itm In sel
        If TypeOf itm Is MailItem Then
            SaveImagesFromMail itm, target
        End If
    Next

    MsgBox "Done. Saved to:" & vbCrLf & target, vbInformation
    Exit Sub

EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "PictureDump"
End Sub

' === CORE ===
Private Sub SaveImagesFromMail(mi As MailItem, ByVal target As String)
    Dim att As Attachment
    For Each att In mi.Attachments
        If IsImageAttachment(att) Then
            Dim base As String: base = SafeFileName(att.FileName)
            If Len(base) = 0 Then base = "image.jpg"
            Dim full As String: full = UniquePath(target, base)
            att.SaveAsFile full
        End If
    Next

    ' Capture base64 inline images from HTML body (e.g., Gmail inline)
    SaveInlineImagesFromHTML mi, target
End Sub

' === PATH BUILDERS ===
Private Function BuildTargetFolderFromSubject(ByVal subject As String) As String
    BuildTargetFolderFromSubject = PicturesPath() & "\" & SafeFolderName(subject)
End Function

' === IMAGE DETECTION ===
Private Function IsImageAttachment(att As Attachment) As Boolean
    On Error GoTo done

    Dim ext As String
    ext = LCase$(Mid$(att.FileName, InStrRev(att.FileName, ".") + 1))
    Select Case ext
        Case "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp", "heic"
            IsImageAttachment = True
            GoTo done
    End Select

    ' PR_ATTACH_MIME_TAG ? "image/..."
    Dim pa As PropertyAccessor
    Set pa = att.PropertyAccessor
    Dim mime As String
    mime = LCase$(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F"))
    If Left$(mime, 6) = "image/" Then IsImageAttachment = True

done:
End Function

' Extract inline base64 images from HTML
Private Sub SaveInlineImagesFromHTML(mi As MailItem, ByVal target As String)
    On Error Resume Next

    Dim html As String: html = mi.HTMLBody
    If Len(html) = 0 Then Exit Sub

    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "data:image/(jpg|jpeg|png|gif|bmp|webp|tif|tiff);base64,([A-Za-z0-9+/=]+)"
    re.Global = True
    re.IgnoreCase = True
    Set matches = re.Execute(html)

    Dim i As Long
    For i = 0 To matches.Count - 1
        Dim ext As String: ext = LCase$(matches(i).SubMatches(0))
        Dim b64 As String: b64 = matches(i).SubMatches(1)
        Dim imgBytes() As Byte: imgBytes = Base64Decode(b64)
        Dim fname As String: fname = UniquePath(target, "inline_" & Format(Now, "yyyymmdd_hhnnss") & "_" & i & "." & ext)
        SaveBytesToFile fname, imgBytes
    Next i
End Sub

Private Function Base64Decode(ByVal base64String As String) As Byte()
    Dim XML As Object, Node As Object
    Set XML = CreateObject("MSXML2.DOMDocument.6.0")
    Set Node = XML.createElement("b64")
    Node.DataType = "bin.base64"
    Node.Text = base64String
    Base64Decode = Node.nodeTypedValue
End Function

Private Sub SaveBytesToFile(ByVal filePath As String, ByRef bytes() As Byte)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'adTypeBinary
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2 'adSaveCreateOverWrite
    stream.Close
End Sub

' === PATH & NAME UTILITIES ===
Private Function PicturesPath() As String
    On Error Resume Next
    Dim p As String
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    p = sh.SpecialFolders("MyPictures")
    If Len(p) = 0 Then
        Dim shellApp As Object, ns As Object
        Set shellApp = CreateObject("Shell.Application")
        Set ns = shellApp.NameSpace(39)
        If Not ns Is Nothing Then p = ns.Self.Path
    End If
    If Len(p) = 0 Then p = Environ$("USERPROFILE") & "\Pictures"
    PicturesPath = p
End Function

Private Sub EnsureFolder(ByVal p As String)
    Dim fso As Object
    Dim parentPath As String
    On Error GoTo oops

    If Len(p) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(p) Then Exit Sub

    parentPath = fso.GetParentFolderName(p)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        EnsureFolder parentPath
    End If

    fso.CreateFolder p
    Exit Sub

oops:
    MsgBox "EnsureFolder failed for path:" & vbCrLf & p & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Folder Create Error"
End Sub

Private Function SafeFolderName(ByVal s As String) As String
    Dim bad As Variant, i As Long

    s = Trim$(s)
    If Len(s) = 0 Then s = "Email_Images"

    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next i

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop

    Dim reserved As Variant
    reserved = Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    For i = LBound(reserved) To UBound(reserved)
        If UCase$(s) = reserved(i) Then s = s & "_"
    Next i

    If Len(s) > 60 Then s = Left$(s, 60)
    SafeFolderName = s
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim nameOnly As String, ext As String, p As Long

    s = Trim$(s)
    If Len(s) = 0 Then s = "image.jpg"

    p = InStrRev(s, ".")
    If p > 0 Then
        nameOnly = Left$(s, p - 1)
        ext = Mid$(s, p)
    Else
        nameOnly = s
        ext = ""
    End If

    nameOnly = SafeFolderName(nameOnly)

    If Len(ext) > 0 Then
        Dim i As Long, cleanExt As String
        For i = 2 To Len(ext)
            If Mid$(ext, i, 1) Like "[A-Za-z0-9]" Then cleanExt = cleanExt & Mid$(ext, i, 1)
        Next i
        If Len(cleanExt) > 0 Then ext = "." & LCase$(cleanExt) Else ext = ""
    End If

    Dim maxLen As Long: maxLen = 120
    If Len(nameOnly) + Len(ext) > maxLen Then
        nameOnly = Left$(nameOnly, maxLen - Len(ext))
    End If

    SafeFileName = nameOnly & ext
End Function

Private Function UniquePath(ByVal folder As String, ByVal baseName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim cleaned As String: cleaned = SafeFileName(baseName)
    Dim nameOnly As String, ext As String, p As Long

    p = InStrRev(cleaned, ".")
    If p > 0 Then
        nameOnly = Left$(cleaned, p - 1)
        ext = Mid$(cleaned, p)
    Else
        nameOnly = cleaned
        ext = ""
    End If

    Dim candidate As String
    Dim n As Long: n = 0
    Do
        If n = 0 Then
            candidate = folder & "\" & nameOnly & ext
        Else
            candidate = folder & "\" & nameOnly & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & n & ext
        End If
        n = n + 1
    Loop While fso.FileExists(candidate)

    UniquePath = candidate
End Function


