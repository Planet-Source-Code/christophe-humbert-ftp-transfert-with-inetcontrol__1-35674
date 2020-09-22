VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSFERT DE FICHIERS"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4440
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame frameTransfertFile 
      Caption         =   "Choix du fichier à transférer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6570
      Left            =   15
      TabIndex        =   2
      Top             =   15
      Width           =   8145
      Begin VB.FileListBox File1 
         Height          =   4965
         Left            =   3555
         TabIndex        =   7
         Top             =   420
         Width           =   4455
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   465
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Height          =   4590
         Left            =   105
         TabIndex        =   5
         Top             =   900
         Width           =   3375
      End
      Begin VB.CommandButton cmdTransfert 
         Caption         =   "&Transférer"
         Height          =   345
         Left            =   3240
         TabIndex        =   4
         Top             =   6120
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   5715
         Width           =   6435
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   7995
         Y1              =   5580
         Y2              =   5580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nom du fichier :"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   5760
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fermer"
      Height          =   300
      Left            =   6600
      TabIndex        =   1
      Top             =   6720
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   7530
      Visible         =   0   'False
      Width           =   5505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdTransfert_Click()
Screen.MousePointer = 11
cmdClose.Enabled = False
cmdTransfert.Enabled = False
'=======================================================================
'TRANSFERT DU FICHIER SUR LE SERVEUR BACKWEB
If Len(Text2.Text) > 6 Then
    UploadFile Inet1, "URL", "NOM_UTILISATEUR", "MOT DE PASSE", Text1.Text, "/MonDossier/" & Text2.Text
Else
    MsgBox "Vous devez sélectionner le fichier à transférer !", vbExclamation + vbOKOnly, "Transfert du fichier."
End If
'=======================================================================
Screen.MousePointer = vbDefault
cmdClose.Enabled = True
cmdTransfert.Enabled = True
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
Dim Val1, Val2 As Integer
    Text1 = Dir1.Path & "\" & File1.FileName
    Val1 = Len(Text1.Text)
    Val2 = Len(Left(Text1.Text, InStrRev(Text1.Text, "\")))
    Text2 = Right(Text1.Text, Val1 - Val2)
End Sub
