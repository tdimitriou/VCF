VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Send(ByVal Recepient As String, _
                        ByVal Sender As String, _
                        ByVal SMTPServer As String, _
                        ByVal SMTPPort As Integer, _
                        ByVal SMTPUsername As String, _
                        ByVal SMTPPassword As String, _
                        ByVal SMTPUseSSL As Boolean, _
                        Optional ByVal Subject As String, _
                        Optional ByVal Body As String, _
                        Optional ByVal Attchments As String, _
                        Optional ByRef ErrorMessage As String) As Boolean
Try:

    On Error GoTo Catch:
    
    Dim CDOMsg      As CDO.Message
    Set CDOMsg = New CDO.Message
        
    CDOMsg.Configuration.Fields(cdoSMTPServer) = SMTPServer
    CDOMsg.Configuration.Fields(cdoSMTPServerPort) = SMTPPort
    CDOMsg.Configuration.Fields(cdoSMTPUseSSL) = SMTPUseSSL
    CDOMsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    CDOMsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    
    CDOMsg.Configuration.Fields(cdoSendUserName) = SMTPUsername
    CDOMsg.Configuration.Fields(cdoSendPassword) = SMTPPassword
    CDOMsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
                
    CDOMsg.Configuration.Fields.Update
    
    CDOMsg.To = Recepient
    CDOMsg.From = Sender
    CDOMsg.Subject = Subject
    CDOMsg.TextBody = Body
    
    Dim Paths
    Paths = Split(Attchments, ";")
    
    Dim Path As String
    Dim Index As Long
    
    For Index = LBound(Paths) To UBound(Paths)
        Path = Trim$(Paths(Index))
        If Len(Path) Then If New_c.FSO.FileExists(Path) Then CDOMsg.AddAttachment Path
    Next
    
    Dim Result As Boolean
    
    CDOMsg.Send
    Result = True
        
Finally:
    
    Send = Result
    Set CDOMsg = Nothing
    Exit Function
          
Catch:

    ErrorMessage = Err.Description
    Result = False
    GoTo Finally
    
End Function

