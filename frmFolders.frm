VERSION 5.00
Begin VB.Form frmFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Destination Folder"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Autor         : Rohullah Habibi
'Date          : Jan 12, 2001
'--------------------------------------------------------------------
'Form          : frmFolders
'Description   : Destination folder selection form.
'                gDestination public variable
'                contains the user selection



Private Sub cmdCancel_Click()
    gDestination = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    gDestination = Dir1.List(Dir1.ListIndex)
    Unload Me
End Sub

Private Sub Drive1_Change()
   On Error GoTo ErrHand
   Dir1.Path = Drive1.Drive
   Dir1.SetFocus
   Exit Sub
ErrHand:
   If Err.Number = 68 Then
      MsgBox "Drive: " & Drive1.Drive & " is unavailable!", vbCritical + vbOKOnly
      Drive1.Drive = "C:"
      Resume Next
   Else
     MsgBox Err.Description, vbCritical + vbOKOnly
   End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub
