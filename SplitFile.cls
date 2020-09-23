VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SplitFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SplitFile"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'--------------------------------------------------------------------
'Autor         : Rohullah Habibi
'Date          : Jan 12, 2001
'--------------------------------------------------------------------
'Class         : SplitFile
'Description   : Splits any type of file
'Dependencies  : Non
'UserInterface : Non, just MsgBox is used in few places.
'Requirements  : Set the following properties:
'                1-SplitThisFile (file name with path to be splitted)
'                2-SplittedFileSize (a numeric value)
'                3-SplittedFileBasedOn (can be BT=byte,KB=kilobyte,MB=megabyte,GB=gigabyte,LN=lines)
'                4-DestinationPath (where you want to save splitted files)
'
'Note          : if the ErrNumber property <> 0, it means an error occured
'                the ErrMessage property contain the error message.
'
'                If you want to show the process in a ProgressBar control
'                then set the ProcessBar property to a ProgressBar control.
'
'                If you want to show the process in a Label control
'                then set the StatusLabel property to a Label control.
'
'                If you want to show splitted files in ListView control
'                then set the FilesListView property to a ListView control.
'
'                if you want to delete the source file
'                (file to be splitted) after splitting
'                call the DeleteSourceFile method.
'
'                if you want to get a confirmation message before
'                deleting the source file,
'                set the ConfirmDelete property to true.
'
'>>>>>>>>>>>>>>> After you set all required properties then
'                call the SplitFile() method to split the file
'--------------------------------------------------------------------

Option Explicit

Public ConfirmDelete As Boolean
Public SplittedFileCount As Integer

Private mvarSplitThisFile As String
Private mvarSplittedFileSize As Double
Private mvarSplittedFileBasedOn As String
Private dblSplitThisFileSize As Double
Private intWriteHandle As Integer
Private intReadHandle As Integer
Private CurrentSplittedFileCount As Integer
Private mvarErrNumber As Integer 'local copy
Private mvarErrMessage As String 'local copy
Private mvarProcessBar As Object 'local copy
Private mvarStatusLabel As Object 'local copy
Private mvarFilesListview As Object 'local copy
Private dblCurrentSplittedFileSize As Double
Private mvarDestinationPath As String 'local copy
Private mvarFileName As String 'local copy

Private Const FLAG2 = "®"
Private Const FLAG1 = "."



Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let DestinationPath(ByVal vData As String)
    mvarDestinationPath = vData
End Property

Public Property Get DestinationPath() As String
    DestinationPath = mvarDestinationPath
End Property


Public Property Set FilesListview(ByVal vData As Object)
    Set mvarFilesListview = vData
End Property


Public Property Get FilesListview() As Object
    Set FilesListview = mvarFilesListview
End Property


Public Property Set StatusLabel(ByVal vData As Object)
    Set mvarStatusLabel = vData
End Property


Public Property Get StatusLabel() As Object
    Set StatusLabel = mvarStatusLabel
End Property


Public Property Set ProcessBar(ByVal vData As Object)
    Set mvarProcessBar = vData
End Property


Public Property Get ProcessBar() As Object
    Set ProcessBar = mvarProcessBar
End Property


Public Property Let ErrMessage(ByVal vData As String)
Attribute ErrMessage.VB_Description = "if an error occured this property will contain the error message"
    mvarErrMessage = vData
End Property


Public Property Get ErrMessage() As String
    ErrMessage = mvarErrMessage
End Property


Public Property Let ErrNumber(ByVal vData As Integer)
Attribute ErrNumber.VB_Description = "if an error occured this property will contain the error number"
    mvarErrNumber = vData
End Property


Public Property Get ErrNumber() As Integer
    ErrNumber = mvarErrNumber
End Property


Public Property Let SplittedFileBasedOn(ByVal vData As String)
Attribute SplittedFileBasedOn.VB_Description = "Can be LINE, MB or KB"
    mvarSplittedFileBasedOn = vData
End Property


Public Property Get SplittedFileBasedOn() As String
    SplittedFileBasedOn = mvarSplittedFileBasedOn
End Property


Public Property Let SplittedFileSize(ByVal vData As Double)
Attribute SplittedFileSize.VB_Description = "Size of the splitted file to be created"
    mvarSplittedFileSize = vData
End Property


Public Property Get SplittedFileSize() As Double
    SplittedFileSize = mvarSplittedFileSize
End Property


Public Property Let SplitThisFile(ByVal vData As String)
Attribute SplitThisFile.VB_Description = "The path and file name to be splitted"
    If Dir(vData) = "" Then
        MsgBox "File: " & vData & " does not exist!", vbCritical + vbOKOnly
        mvarSplitThisFile = ""
        dblSplitThisFileSize = 0
    Else
        mvarSplitThisFile = vData
        dblSplitThisFileSize = FileLen(vData)
        FileName = GetFileName(vData)
    End If
End Property


Public Property Get SplitThisFile() As String
    SplitThisFile = mvarSplitThisFile
End Property


Public Function SplitFile()

    If Dir(SplitThisFile) = "" Then
        ErrMessage = "File: " & SplitThisFile & " does not exist!"
        ErrNumber = -1
        Exit Function
    End If
    
    If SplittedFileSize = 0 Then
        ErrMessage = "Please specify file size the file to be splitted!"
        ErrNumber = -1
        Exit Function
    End If
    
    On Error GoTo Errorhand
    
    CurrentSplittedFileCount = 0
    SplittedFileCount = 0
    
    
    If TypeName(FilesListview) = "ListView" Then
        FilesListview.ListItems.Clear
    End If
    
    If SplittedFileBasedOn = "LN" Then
        ErrNumber = SplitIntoTextFiles()
    Else
        ErrNumber = SplitIntoBinaryFiles()
    End If
    If ErrNumber <> 0 Then
        If CurrentSplittedFileCount > 0 Then
           DeletePartialSplit
        End If
    End If
    
    
    Exit Function
    
Errorhand:
        
        If CurrentSplittedFileCount > 0 Then
            DeletePartialSplit
        End If
        
        ErrMessage = "Error: " & Err.Number & " - " & Err.Description
        ErrNumber = -1
End Function


Private Function SplitIntoBinaryFiles() As Integer
    
    Dim i As Integer
    Dim dblSize As Double
        
    Dim ByteArray() As Byte
    
    Select Case SplittedFileBasedOn
        Case "BT"
            dblSize = SplittedFileSize
        Case "KB"
            dblSize = SplittedFileSize * 1024
        Case "MB"
            dblSize = (SplittedFileSize * 1024) * 1024
        Case "GB"
            dblSize = ((SplittedFileSize * 1024) * 1024) * 1024
    End Select
    
    If dblSplitThisFileSize < dblSize Then
       ErrMessage = "The file you want to split is smaller than the size you specified!!!"
       SplitIntoBinaryFiles = -1
       Exit Function
    End If
    
    SplittedFileCount = Abs(Int(-Format(dblSplitThisFileSize / dblSize, "#########0.00")))
    If LCase(Left(DestinationPath, 2)) = "a:" Then
        If Format((dblSize / 1024) / 1024, "#.00") > "1.38" Then
            ErrMessage = "The split size you specified is greater than the " & _
                         Chr(13) & "diskette storage size!"
            SplitIntoBinaryFiles = -1
            Exit Function
        End If
        If MsgBox("To split " & SplitThisFile & Chr(13) & _
            " to drive A:, you need " & Trim(Str(SplittedFileCount)) & _
            "  1.44MB formatted diskettes." & Chr(13) & Chr(13) & _
            "If you have " & Trim(Str(SplittedFileCount)) & _
            "  1.44MB formatted diskettes, click Yes otherwise click No ?", vbYesNo + vbQuestion) = vbNo Then
            SplitIntoBinaryFiles = -1
            ErrMessage = "Process Canceled."
            Exit Function
        End If
    End If
    ReDim ByteArray(1 To dblSize)
    
    intReadHandle = FreeFile
    Open SplitThisFile For Binary Access Read As #intReadHandle
    
    If TypeName(ProcessBar) = "ProgressBar" Then
       ProcessBar.Max = SplittedFileCount
    End If
    
    For i = 1 To SplittedFileCount
        DoEvents
          
        dblCurrentSplittedFileSize = UBound(ByteArray, 1)
    
        If OpenNewFile() <> 0 Then
            Exit For
        End If
        
        Get #intReadHandle, , ByteArray
        Put #intWriteHandle, , ByteArray
        
        If dblSplitThisFileSize - (i * UBound(ByteArray, 1)) > 0 And _
           dblSplitThisFileSize - (i * UBound(ByteArray, 1)) < UBound(ByteArray, 1) Then
           ReDim ByteArray(1 To dblSplitThisFileSize - (i * UBound(ByteArray, 1)))
        End If
    Next
    
    ReDim ByteArray(0)
    SplitIntoBinaryFiles = ErrNumber
End Function


Private Function SplitIntoTextFiles() As Integer
    
    Dim dblLineCounter As Double
    Dim strOneLine As String
   
    intReadHandle = FreeFile
   
    Open SplitThisFile For Input As #intReadHandle
    
    Line Input #intReadHandle, strOneLine
    
    SplittedFileCount = Format(dblSplitThisFileSize / (Len(Trim(strOneLine)) + 2) / SplittedFileSize, "#########0.0")
    If LCase(Left(DestinationPath, 2)) = "a:" Then
        If MsgBox("To split " & SplitThisFile & Chr(13) & _
            " to drive A:, you need " & Trim(Str(SplittedFileCount)) & _
            "  1.44MB formatted diskettes." & Chr(13) & Chr(13) & _
            "If you have " & Trim(Str(SplittedFileCount)) & _
            "  1.44MB formatted diskettes, click Yes otherwise click No ?", vbYesNo + vbQuestion) = vbNo Then
            SplitIntoTextFiles = -1
            ErrMessage = "Process Canceled."
            Exit Function
        End If
    End If
    
    dblCurrentSplittedFileSize = Format(((Len(strOneLine) + 2) * SplittedFileSize) / 1024, "########")
    If TypeName(ProcessBar) = "ProgressBar" Then
       ProcessBar.Max = SplittedFileCount
    End If
    Seek #intReadHandle, 1
    
    dblCurrentSplittedFileSize = (Len(strOneLine) + 2) * SplittedFileSize
    
    If OpenNewFile() <> 0 Then
        SplitIntoTextFiles = ErrNumber
        Exit Function
    End If
        
    Do While Not EOF(intReadHandle)  ' Loop until end of file.
       DoEvents
       
       
       If SplittedFileSize = dblLineCounter Then
           dblLineCounter = 0
           
           dblCurrentSplittedFileSize = LOF(intWriteHandle)
           
           If OpenNewFile <> 0 Then
              Exit Do
           End If
       End If
       
       Line Input #intReadHandle, strOneLine
       Print #intWriteHandle, strOneLine
       
       dblLineCounter = dblLineCounter + 1
       
    Loop
    SplitIntoTextFiles = ErrNumber
End Function

Private Function OpenNewFile() As Integer
        
        On Error GoTo ErrHand
        Dim strFile As String
        
        Close #intWriteHandle
        
        CurrentSplittedFileCount = CurrentSplittedFileCount + 1
        
        strFile = DestinationPath & FileName & FLAG1 & SplittedFileCount & FLAG2 & CurrentSplittedFileCount
        
        intWriteHandle = FreeFile
        
        If LCase(Left(DestinationPath, 2)) = "a:" Then
            If MsgBox("Please insert disk #" & CurrentSplittedFileCount & " in drive A:.", vbInformation + vbOKCancel) = vbCancel Then
               ErrNumber = -1
               ErrMessage = "Process canceled by the user."
               OpenNewFile = ErrNumber
               Exit Function
            Else
                Dir ("a:")
            End If
        End If
           
        If SplittedFileBasedOn = "LN" Then
            Open strFile For Output As #intWriteHandle
        Else
            Open strFile For Binary Access Write As #intWriteHandle
        End If
        
        
        If TypeName(ProcessBar) = "ProgressBar" Then
           ProcessBar.Visible = True
           If CurrentSplittedFileCount <= ProcessBar.Max Then
                ProcessBar.Value = CurrentSplittedFileCount
           End If
        End If
        If TypeName(StatusLabel) = "Label" Then
            If Len("Creating File: " & strFile) > 80 Then
                StatusLabel.Caption = "Creating File: " & "..." & Right$(strFile, 51)
            Else
                StatusLabel.Caption = "Creating File: " & strFile
            End If
        End If
        If TypeName(FilesListview) = "ListView" Then
           AddListViewItem FilesListview, strFile, dblCurrentSplittedFileSize
        End If
        
        
        DoEvents
        
        OpenNewFile = 0
        Exit Function
ErrHand:
      If Err.Number = 52 Then 'bad file name or number
         If LCase(Left(DestinationPath, 2)) = "a:" Then
            If MsgBox("Please insert disk #" & CurrentSplittedFileCount & " in drive A:.", vbInformation + vbOKCancel) = vbCancel Then
               ErrNumber = -1
               ErrMessage = "Process canceled by the user."
               OpenNewFile = ErrNumber
               Exit Function
            Else
               Resume
            End If
         Else
            ErrMessage = Err.Description
            ErrNumber = Err.Number
            OpenNewFile = Err.Number
         End If
    Else
        ErrMessage = Err.Description
        ErrNumber = Err.Number
        OpenNewFile = Err.Number
   End If
End Function


Private Sub Class_Initialize()
    dblSplitThisFileSize = 0
    CurrentSplittedFileCount = 0
    dblCurrentSplittedFileSize = 0
    SplittedFileCount = 0
    intReadHandle = 0
    intWriteHandle = 0
    ErrMessage = ""
    ErrNumber = 0
    ConfirmDelete = True
    DestinationPath = ""
    FileName = ""
End Sub

Private Sub Class_Terminate()
    If TypeName(ProcessBar) = "ProgressBar" Then
       ProcessBar.Value = 0
       ProcessBar.Visible = False
       Set ProcessBar = Nothing
    End If
    If TypeName(StatusLabel) = "Label" Then
        StatusLabel.Caption = ""
        Set StatusLabel = Nothing
    End If
    If TypeName(FilesListview) = "ListView" Then
        'FilesListview.ListItems.Clear
        Set FilesListview = Nothing
    End If
    
    Close #intWriteHandle
    Close #intReadHandle
End Sub

Private Function DeletePartialSplit()
    Dim i As Integer
    Close #intWriteHandle
    Close #intReadHandle
    If LCase(Left(DestinationPath, 2)) <> "a:" Then
        For i = 1 To CurrentSplittedFileCount
            If Dir(DestinationPath & FileName & FLAG1 & SplittedFileCount & FLAG2 & Trim(Str(i))) <> "" Then
                Kill (DestinationPath & FileName & FLAG1 & SplittedFileCount & FLAG2 & Trim(Str(i)))
            End If
        Next
    End If
End Function

Friend Sub AddListViewItem(lv As ListView, strCurrentSplittedFile As String, dblSplittedFilesize As Double)
    Dim itmx As ListItem
    
    Set itmx = lv.ListItems.Add(, , strCurrentSplittedFile, , 1)
    If SplittedFileBasedOn = "" Then
        SplittedFileBasedOn = "KB"
    End If
    Select Case SplittedFileBasedOn
        Case "BT"
            itmx.SubItems(1) = Format(dblSplittedFilesize, "###########") & "Bytes"
        Case "KB", "LN"
            itmx.SubItems(1) = Format(dblSplittedFilesize / 1024, "##########0.00") & "KB"
        Case "MB"
            If (dblSplittedFilesize / 1024) Mod 1024 > 0 Then
               If (dblSplittedFilesize / 1024) / 1024 < 1 Then
                   itmx.SubItems(1) = Format(dblSplittedFilesize / 1024, "##########0.00") & "KB"
               Else
                    itmx.SubItems(1) = Format((dblSplittedFilesize / 1024) / 1024, "##########0.00") & "MB"
               End If
            Else
                itmx.SubItems(1) = Format((dblSplittedFilesize / 1024) / 1024, "##########0.00") & "MB"
            End If
        Case "GB"
            If ((dblSplittedFilesize / 1024) / 1024) Mod 1024 > 0 Then
               If (dblSplittedFilesize / 1024) / 1024 < 1000 Then
                   itmx.SubItems(1) = Format((dblSplittedFilesize / 1024) / 1024, "##########0.00") & "MB"
               Else
                    itmx.SubItems(1) = Format(((dblSplittedFilesize / 1024) / 1024) / 1024, "##########0.00") & "GB"
               End If
            Else
                itmx.SubItems(1) = Format(((dblSplittedFilesize / 1024) / 1024) / 1024, "##########") & "GB"
            End If
        
    End Select
    itmx.EnsureVisible
    Set itmx = Nothing
End Sub

Public Function DeleteSourceFile()
    If ErrNumber = 0 Then
        If Dir(SplitThisFile) <> "" Then
            If ConfirmDelete Then
                If MsgBox("The file: " & Chr(13) & _
                    SplitThisFile & Chr(13) & " has been spliited successfully," & _
                    " do you wish to delete it?", vbQuestion + vbOKCancel) = vbCancel Then
                    Exit Function
                End If
            End If
            Close (intReadHandle)
            Close (intWriteHandle)
            Kill (SplitThisFile)
        End If
    End If
End Function


'this function extracts file name from a path with filename
Friend Function GetFileName(strFileWithPath As String) As String
    If InStr(1, strFileWithPath, "\") = 0 Then
        GetFileName = strFileWithPath
        Exit Function
    End If
    Do
        strFileWithPath = Mid$(strFileWithPath, _
                          IIf(InStr(1, strFileWithPath, "\") > 0, _
                              InStr(1, strFileWithPath, "\") + 1, 1))
    Loop While InStr(1, strFileWithPath, "\") <> 0
    GetFileName = strFileWithPath
End Function


