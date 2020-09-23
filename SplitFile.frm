VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmsplit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CuteSplit 2.0"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":0000
            Key             =   "FILE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":031A
            Key             =   "EXIT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":0424
            Key             =   "ABOUT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":0876
            Key             =   "SPLITF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":14C8
            Key             =   "JOIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":211A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":2D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":39BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":4610
            Key             =   "SPLIT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SplitFile.frx":5262
            Key             =   "TOOLS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5653
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   9975
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   5520
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   5633
      Index           =   0
      Left            =   1560
      ScaleHeight     =   5580
      ScaleWidth      =   6525
      TabIndex        =   1
      Top             =   0
      Width           =   6585
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -120
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2715
         Left            =   0
         TabIndex        =   16
         Top             =   2880
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   4789
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Splitted Files"
            Object.Width           =   8467
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "Size"
            Text            =   "Size"
            Object.Width           =   2682
         EndProperty
      End
      Begin VB.Frame FraAbout 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   6495
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            Picture         =   "SplitFile.frx":5EB4
            ScaleHeight     =   1095
            ScaleWidth      =   1215
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   $"SplitFile.frx":9AA2
            Height          =   855
            Left            =   120
            TabIndex        =   34
            Top             =   4200
            Width           =   5415
         End
         Begin VB.Label Label5 
            Caption         =   "Copyright (c) 2001 R. Habibi."
            Height          =   255
            Left            =   3360
            TabIndex        =   33
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "CuteSplit 2.0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3240
            TabIndex        =   32
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame FraOptions 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   0
         TabIndex        =   25
         Top             =   480
         Width           =   6495
         Begin VB.CheckBox chkDelSplitConfirm 
            Caption         =   "Confirm when deleting all splitted files"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   1800
            Width           =   3495
         End
         Begin VB.CheckBox chkDelSplit 
            Caption         =   "Delete splitted files after successful join"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CheckBox chkDelSourceConfirm 
            Caption         =   "Confirm when deleting the source file"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   1440
            Width           =   3975
         End
         Begin VB.CheckBox chkDelSource 
            Caption         =   "Delete source file after it is splitted successfully"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   720
            Width           =   5175
         End
      End
      Begin VB.Frame FraSplit 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6255
         Begin VB.CommandButton cmdSplitDest 
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Height          =   300
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1680
            Width           =   300
         End
         Begin VB.TextBox txtSplitDest 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1680
            Width           =   4575
         End
         Begin VB.CommandButton cmdSplit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Split"
            Height          =   375
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox CboSize 
            Height          =   315
            ItemData        =   "SplitFile.frx":9BC1
            Left            =   2520
            List            =   "SplitFile.frx":9BD4
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtsplit 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   4575
         End
         Begin VB.CommandButton cmdBrowse 
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Height          =   300
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Splitted File Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "File to be splitted:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FraJoin 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Width           =   6375
         Begin VB.CommandButton cmdJoinDest 
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Height          =   300
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1080
            Width           =   300
         End
         Begin VB.TextBox txtJoinDest 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1080
            Width           =   4575
         End
         Begin VB.CommandButton cmdJoin 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Join"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdBrowseJoin 
            BackColor       =   &H00C0C0C0&
            Caption         =   "..."
            Height          =   300
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Width           =   300
         End
         Begin VB.TextBox txtJoin 
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination:"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Locate the first splitted file:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Label lblcaption 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6615
      End
      Begin VB.Label lblstatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmsplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Autor         : Rohullah Habibi
'Date          : Jan 12, 2001
'--------------------------------------------------------------------
'Application   : Split/Join Files
'Description   : Split/join any type of file
'                pure VB code no Dll or Activex and very fast.
'Dependencies  : SplitFile Class and JoinFile Class




Private Sub cmdBrowse_Click()
   'get the file to be splitted
   txtsplit.Text = LCase(LocateFile())
   'set destination path same as source file if it is blank
   If txtSplitDest.Text = "" Then txtSplitDest.Text = GetFilePath(txtsplit.Text)
   'go to size text box
   Text2.SetFocus
End Sub

Private Sub cmdBrowseJoin_Click()
    
    Dim strFile As String
    'get the first splitted file to create the joined file
    strFile = LCase(LocateFile())
    If strFile <> "" Then
        
        'instantiate a new instance of the join class
        Dim JoinIt As New JoinFile
        
        'set some properties of the join object
        
        JoinIt.JoinThisFile = strFile
        'if you want to display the list of splitted files
        'set this property to a listview object
        Set JoinIt.FilesListview = ListView2
        
        'locate the first and all splitted files
        JoinIt.LocateSplitedFiles
        
        'if the ErrNumber property <> 0 then an error occured in the
        'Join class and the ErrMeesage property has the last error message.
        If JoinIt.ErrNumber <> 0 Then
            MsgBox JoinIt.ErrMessage, vbCritical + vbOKOnly
            cmdJoin.Enabled = False
            txtJoin.Text = ""
            txtJoinDest.Text = ""
        Else
            'if no error so far, then the user have all splitted files
            cmdJoin.Enabled = True
            txtJoin.Text = strFile
            If txtJoinDest.Text = "" And Left(strFile, 2) <> "a:" Then txtJoinDest.Text = GetFilePath(txtJoin.Text)
        End If
        'release the instance
        Set JoinIt = Nothing
    End If
End Sub


Private Sub cmdJoin_Click()
    'check if the the user selected a splitted file to be joined
    If txtJoin.Text = "" Then
        MsgBox "Please locate the first splitted file!", vbCritical + vbOKOnly
        Exit Sub
    End If
    'check if the the user selected the destination
    If txtJoinDest.Text = "" Then
        MsgBox "Please specify destination for the joined file!", vbCritical + vbOKOnly
        Exit Sub
    End If
        
        'create an instance of the join class and set some properties
        Dim JoinIt As New JoinFile
        JoinIt.JoinThisFile = txtJoin.Text
        
        JoinIt.DestinationPath = txtJoinDest.Text
        
        'if you want to show the process in the ProgressBar, then set this
        'property to a ProgressBar control
        Set JoinIt.ProcessBar = ProgressBar1
        'if you want to show the process in the label control, then
        'set this property to a label control
        Set JoinIt.StatusLabel = lblstatus
        'if you want to show all splitted file in a listview, then set this
        'property to a ListView control
        Set JoinIt.FilesListview = ListView2
        
        'once again check to make sure the user has all the right splitted files
        JoinIt.LocateSplitedFiles
        
        'if no error in the Join object then proceed with
        'joining all splitted files
        If JoinIt.ErrNumber <> 0 Then
            MsgBox JoinIt.ErrMessage
        Else
            
            Screen.MousePointer = vbHourglass
            JoinIt.JoinFiles
            Screen.MousePointer = vbDefault
        
            If JoinIt.ErrNumber <> 0 Then
               MsgBox "Error: " & JoinIt.ErrNumber & " : " & _
               JoinIt.ErrMessage, vbCritical + vbOKOnly
            Else
              MsgBox "The file: " & JoinIt.DestinationPath & JoinIt.FileName & Chr(13) & _
              " has been created from all splitted files successfully.", vbInformation + vbOKOnly
            
                'if the user wants to delete all splitted files after join, then
                'get red of them.
                'Note: to set this option click on "Option" icon.
                If chkDelSplit.Value Then
                    JoinIt.ConfirmDelete = chkDelSplitConfirm.Value
                    JoinIt.DeleteSplittedFiles
                End If
            End If
        End If
        'we are done, release the instance
        Set JoinIt = Nothing

End Sub

Private Sub cmdJoinDest_Click()
    'pop the destination folder selection form and let
    'the user pick the destination where he/she can save the joined file
    frmFolders.Show vbModal, Me
    If gDestination <> "" Then
        txtJoinDest.Text = LCase(IIf(Right(gDestination, 1) <> "\", gDestination & "\", gDestination))
    End If
End Sub

Private Sub cmdSplit_Click()
    Dim strSplitSize As String
    'check if the the user selected a file to be splitted
    If txtsplit.Text = "" Then
        MsgBox "Please select a file to be splitted.", vbCritical + vbOKOnly
        Exit Sub
    End If
    'check if the the user selected the destination for splitted files
    If txtSplitDest.Text = "" Then
        MsgBox "Please specify destination for the splitted files!", vbCritical + vbOKOnly
        Exit Sub
    End If
   
   'check if the the user enetred a size for the file to splitted into.
   If Text2.Text = "" Then
        MsgBox "Invalid splitted file size value!", vbCritical + vbOKOnly
        Text2.SetFocus
        Exit Sub
    End If
    
    
    Select Case CboSize.ListIndex
        Case 0
            strSplitSize = "BT"
        Case 1
            strSplitSize = "KB"
        Case 2
            strSplitSize = "MB"
        Case 3
            strSplitSize = "GB"
        Case 4
            strSplitSize = "LN"
    End Select
    
    'ok, the user entered all required inputs, so lets create the
    'Split object and set some properties
    Dim SplitIt As New SplitFile
    
    SplitIt.SplitThisFile = txtsplit.Text
    
    SplitIt.SplittedFileSize = Val(Text2.Text)
    
    SplitIt.DestinationPath = txtSplitDest.Text
    
    SplitIt.SplittedFileBasedOn = strSplitSize

    'if you want to show the process in the ProgressBar, then set this
    'property to a ProgressBar control
    Set SplitIt.ProcessBar = ProgressBar1
    
    'if you want to show the process in the Label control, then set this
    'property to a Label control
    Set SplitIt.StatusLabel = lblstatus
    
    'if you want to show the list of splitted files in the ListView control, then
    'set this property to a ListView control
    Set SplitIt.FilesListview = ListView2
    
    
    'Ok, we are all set lets get the job done.
    Screen.MousePointer = vbHourglass
    SplitIt.SplitFile
    Screen.MousePointer = vbDefault
    
    'if the ErrNumber property is <> 0 then we have a problem
    'somthing went wrong in the Split object, to find out,
    'check the ErrMessage property.
    If SplitIt.ErrNumber <> 0 Then
        If SplitIt.ErrNumber <> -1 Then
            MsgBox "Error: " & SplitIt.ErrNumber & " : " & _
                SplitIt.ErrMessage, vbCritical + vbOKOnly
        Else
            MsgBox SplitIt.ErrMessage, vbCritical + vbOKOnly
        End If
    Else
        'if we are here, it meane the job is done, so lets make the user
        'happy and display the process complete message.
        MsgBox "The file " & txtsplit.Text & Chr(13) & _
                " has been splitted into " & _
                SplitIt.SplittedFileCount & _
                " files each containing  " & _
                Trim(Text2.Text) & " " & CboSize.List(CboSize.ListIndex) & Chr(13) & _
                "The last file may contain less then " & _
                Trim(Text2.Text) & " " & CboSize.List(CboSize.ListIndex), vbInformation + vbOKOnly
        
        'if the user wants to delete the source file (the file that is
        'supposed to be splitted) after split, then
        'get red of it.
        'Note: to set this option click on "Option" icon.
        If chkDelSource.Value Then
            SplitIt.ConfirmDelete = chkDelSourceConfirm.Value
            SplitIt.DeleteSourceFile
        End If
    
    End If
    'release the instance
    Set SplitIt = Nothing
    
End Sub


Private Sub cmdSplitDest_Click()
    'pop the destination folder selection form and let
    'the user pick the destination where he/she can save all splitted files
    frmFolders.Show vbModal, Me
    If gDestination <> "" Then
        txtSplitDest.Text = LCase(IIf(Right(gDestination, 1) <> "\", gDestination & "\", gDestination))
    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'set default split size to Megabyte
    CboSize.ListIndex = 2
    
    'Add control icons on the ListView
    ListView1.ListItems.Add 1, "SPLIT", "Split Files", "SPLIT"
    ListView1.ListItems.Add 2, "JOIN", "Join Files", "JOIN"
    ListView1.ListItems.Add 3, "TOOLS", "Options", "TOOLS"
    ListView1.ListItems.Add 4, "ABOUT", "About", "ABOUT"
    ListView1.ListItems.Add 5, "EXIT", "Exit", "EXIT"
    
    Picture1.Item(0).Width = 6585
    Picture1.Item(0).Height = 5653
    Picture1.Item(0).Left = 1595
    
    'set the caption for the default right panel window
    lblcaption.Caption = "  Split Files"
    SetFrame "SPLIT"
    
    'get setting from windows registry for the "Option" screen
    'Note: to see these options click on "Option" icon
    chkDelSource.Value = GetRegistryValue(App.EXEName, "Options", "DelSource", 0)
    chkDelSplit.Value = GetRegistryValue(App.EXEName, "Options", "DelSplit", 0)
    chkDelSplitConfirm.Value = GetRegistryValue(App.EXEName, "Options", "DelSplitConfirm", 0)
    chkDelSourceConfirm.Value = GetRegistryValue(App.EXEName, "Options", "DelSourceConfirm", 0)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Save setting to windows registry for the "Option" screen
    'Note: to see these options click on "Option" icon
    SaveRegistryValue App.EXEName, "Options", "DelSource", chkDelSource.Value
    SaveRegistryValue App.EXEName, "Options", "DelSplit", chkDelSplit.Value
    SaveRegistryValue App.EXEName, "Options", "DelSplitConfirm", chkDelSplitConfirm.Value
    SaveRegistryValue App.EXEName, "Options", "DelSourceConfirm", chkDelSourceConfirm.Value
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'allow only numeric value in Size textbox
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub


Private Sub ListView1_Click()

    'do nothing if the user clicks on disabled icon
    If ListView1.SelectedItem.Ghosted Then Exit Sub
    
        
    ListView1.Refresh
    Dim itemx As ListItem
    
    
    Set itemx = ListView1.SelectedItem
    
    'do nothing if the user clicks on an area other than icons in the listview
    If Not itemx.Selected Then Exit Sub
    
    'set the label caption
    If itemx.Key <> "EXIT" Then
        lblcaption.Caption = "  " & itemx.Text
    End If
    
    Dim strKey As String
    strKey = itemx.Key
    Set itemx = Nothing
    ListView1.Tag = ""
    'show/hide frames according to user selection
    SetFrame strKey
End Sub

'This function shows/hides frames based on user selection in the listview
Function SetFrame(Key As String)
    If Key <> "EXIT" Then
        FraSplit.Visible = False
        FraJoin.Visible = False
        FraOptions.Visible = False
        FraAbout.Visible = False
        
        If ListView2.ListItems.Count > 0 Then
            ListView2.ListItems.Clear
        End If
        ListView2.Visible = True
    End If
    Select Case Key
        Case "SPLIT"
             FraSplit.Visible = True
        Case "JOIN"
             FraJoin.Visible = True
        Case "TOOLS"
            FraOptions.Visible = True
            ListView2.Visible = False
        Case "ABOUT"
            ListView2.Visible = False
            FraAbout.Visible = True
        Case "EXIT"
        
            If MsgBox("Do you wish to close this application?", vbQuestion + vbYesNo) = vbYes Then
                Unload Me
            End If
    End Select
End Function


'This function pops the file selection window of CommonDialog control.
Function LocateFile() As String

      With CmDlg
       .Filter = "*.*"
       .InitDir = App.Path
       .FileName = ""
       .ShowOpen
       If Not .CancelError And .FileName <> "" And Dir(.FileName) <> "" Then
          LocateFile = LCase(.FileName)
       Else
          LocateFile = ""
       End If
   End With
End Function

