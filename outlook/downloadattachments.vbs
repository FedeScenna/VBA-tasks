Public Sub SaveOLFolderAttachments()
      
  ' Ask the user to select an Outlook folder to process
  Dim olPurgeFolder As Outlook.MAPIFolder
  Set olPurgeFolder = Outlook.GetNamespace("MAPI").PickFolder
  If olPurgeFolder Is Nothing Then Exit Sub
   
  
  ' Ask the user to select a file system folder for saving the attachments
  Dim oShell As Object
  Set oShell = CreateObject("Shell.Application")
  Dim fsSaveFolder As Object
  Set fsSaveFolder = oShell.BrowseForFolder(0, "Please Select a Save Folder:", 1)
  If fsSaveFolder Is Nothing Then Exit Sub
  ' Note:  BrowseForFolder doesn't add a trailing slash



  ' Iteration variables
    Dim msg As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim sSavePathFS As String
    Dim sDelAtts
    Dim msg_title As String
    Dim fileFormat As String
  For Each msg In olPurgeFolder.Items
    
    sDelAtts = ""

    If msg.Attachments.Count > 0 Then

      ' This While loop is controlled via the .Delete method
      ' which will decrement msg.Attachments.Count by one each time.
      Dim adj
      For Each att In msg.Attachments
      

        ' Save the file
        msg_title = msg.Subject
        fileFormat = Right(msg.Attachments(1).filename, 5)
        sSavePathFS = fsSaveFolder.Self.Path & "\" & msg_title & fileFormat
        'sSavePathFS = fsSaveFolder.Self.Path & "\" & msg.Attachments(1).filename
        msg.Attachments(1).SaveAsFile sSavePathFS

      Next

    End If

  Next

End Sub