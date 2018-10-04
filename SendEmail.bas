Sub SendEmail()
Dim OutlookApp As Outlook.Application
Dim OutlookItem As Outlook.MailItem
Set OutlookApp = New Outlook.Application
Set OutlookItem = OutlookApp.CreateItem(olMailItem)

Receiver = [AN1].Value
CCReceiver = [AN5].Value

SubjectText = [AN2].Value
AttachedObject = [AN4].Value

On Error GoTo SendEmail_Error
    With OutlookItem
    .To = Receiver
    .CC = CCReceiver
    .Subject = SubjectText
    .BodyFormat = olFormatHTML
    .HTMLBody = "<HTML><p1>Dear all,<p2><BODY><br>Please check the Report in the attachment. <p3><BODY><br>Thanks.<br></BODY></HTML>"
    If AttachedObject <> "" Then
    .Attachments.Add AttachedObject
    End If
    
    .Display
    End With
SendEmail_Exit:
    Exit Sub
SendEmail_Error:
    MsgBox Err.Description
    Resume SendEmail_Exit

End Sub
