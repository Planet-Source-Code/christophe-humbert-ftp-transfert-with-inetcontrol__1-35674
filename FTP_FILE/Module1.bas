Attribute VB_Name = "Module1"
Public FExists As Boolean
Public StatFileTransfert As Boolean

'===================================================================
'Fonction permettant le transfert du fichier.
Public Function UploadFile(InetControl As Inet, ByVal strURL As String, ByVal strNomUtilisateur As String, ByVal strMotDePasse As String, ByVal strFichierLocal As String, ByVal strFichierDistant As String) As Boolean
    On Error GoTo ErrHandle_UploadFile
    
    If InetControl.StillExecuting Then GoTo ErrHandle_UploadFile
    
    With InetControl
        .Cancel
        .Protocol = icFTP
        .URL = strURL
        .UserName = strNomUtilisateur
        .Password = strMotDePasse
    End With
        
    InetControl.Execute , "PUT " & Chr(34) & strFichierLocal & Chr(34) & " " & Chr(34) & strFichierDistant & Chr(34)
        
    Do While InetControl.StillExecuting
        DoEvents
    Loop
            
    UploadFile = True
    StatFileTransfert = True
    MsgBox "Le fichier a été transferé sur le serveur", vbInformation + vbOKOnly, "Transfert"
      

Exit Function
ErrHandle_UploadFile:
            UploadFile = False
            MsgBox "Le fichier n'a pas été transferé sur le serveur !", vbExclamation + vbOKOnly, "Transfert"
            Exit Function
End Function
'===================================================================

Public Function FileExists(ByVal FileName As String)
Dim Exists As Integer
   
On Local Error Resume Next

Exists = Len(Dir(FileName$))

On Local Error GoTo 0
If Exists = 0 Then
    FileExists = False
    FExists = False
Else
    FileExists = True
    FExists = True
End If

End Function

