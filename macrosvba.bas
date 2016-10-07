'references
'vba
'ole automation
'ms office 14 library
'microsoft scripting runtime

Private Sub Application_NewMail()
SaveAttachments
End Sub
Sub CreateDir(strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Sub
Public Sub SaveAttachments()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String
Dim fs As FileSystemObject

' Get the path to your My Documents folder
strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
'On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

' Set the Attachment folder.
strFolderpath = "E:\Outlook\Attachments\"

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objMsg In objSelection

    ' This code only strips attachments from mail items.
    ' If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
    strDeletedFiles = ""

    If lngCount > 0 Then

        ' We need to use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.
        Set fs = New FileSystemObject

        For i = lngCount To 1 Step -1


Dim fn As String
fn = objAttachments.Item(i).FileName
Dim extlen As Integer
extlen = InStrRev(fn, ".")
Dim ext As String
ext = Mid(fn, extlen + 1)
Dim filenamenoxet As String
filenamenoext = Mid(fn, 1, extlen - 1)



            ' Save attachment before deleting from item.
            ' Get the file name.
           ' strFile = Left(fn, Len(fn) - 4) + "_" + Right("00" + Trim(Str$(Day(Now))), 2) + "_" + Right("00" + Trim(Str$(Month(Now))), 2) + "_" + Right("0000" + Trim(Str$(Year(Now))), 4) + "_" + Right("00" + Trim(Str$(Hour(Now))), 2) + "_" + Right("00" + Trim(Str$(Minute(Now))), 2) + "_" + Right("00" + Trim(Str$(Second(Now))), 2) + Right((fn), 4)

Dim fullpath As String
fullpath = strFolderpath + "\" + CStr(Year(Now)) + "\" + CStr(Month(Now)) + "\" + CStr(Day(Now))
CreateDir (fullpath)

strFile = fullpath + "\" + filenamenoext + "." + ext

            ' Combine with the path to the Temp folder.
          '  strFile = strFolderpath & strFile

            ' Save the attachment as a file.
            objAttachments.Item(i).SaveAsFile strFile

            ' Delete the attachment.
            objAttachments.Item(i).Delete

            'write the save as path to a string to add to the message
            'check for html and use html tags in link
            If objMsg.BodyFormat <> olFormatHTML Then
                strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
            Else
                strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
                strFile & "'>" & strFile & "</a>"
            End If

            'Use the MsgBox command to troubleshoot. Remove it from the final code.
            'MsgBox strDeletedFiles

        Next i

        ' Adds the filename string to the message body and save it
        ' Check for HTML body
        If objMsg.BodyFormat <> olFormatHTML Then
            objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
        Else
            objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
        End If

        objMsg.Save
    End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub

