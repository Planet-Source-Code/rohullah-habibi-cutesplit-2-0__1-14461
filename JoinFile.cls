VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JoinFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
'--------------------------------------------------------------------
'Autor         : Rohullah Habibi
'Date          : Jan 12, 2001
'--------------------------------------------------------------------
'Class         : JoinFiles
'Description   : Joins splitted files splitted by the SplitFile object
'Dependencies  : SplitFile Class
'UserInterface : No, just MsgBox is used in few places.
'Requirements  : Make sure the FLAG1 and FLAG2 properties match with
'                SplitFile class properties.
'                Set the following properties:
'                1-JoinThisFile (the first splitted file)
'                2-DestinationPath (where you want to save the joined file)
'
'Note          : if the ErrNumber property <> 0, it means an error occured
'                the ErrMessage property contain the error message.
'
'                After you set up all properties,
'                call the LocateSplittedFiles() method to find out if
'                the user selected the right splitted file.
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
'                if you want to delete all splitted files after join,
'                call the DeleteSplittedfiles method.
'
'                if you want to get a confirmation message before
'                deleting all splitted files,
'                set the ConfirmDelete property to true.
'
'>>>>>>>>>>>>>>> After you set all required properties then
'                call the JoinFiles() method to split the file
'
'--------------------------------------------------------------------

Option Explicit

Public JoinFileCount As Integer
Public JoinedFileName As String
Public ConfirmDelete As Boolean


Private splitObject As New SplitFile 'Just to call some friend procedures
Private intWriteHandle As Integer
Private intReadHandle As Integer
Private mvarErrMessage As String 'local copy
Private mvarErrNumber As Integer 'local copy
Private mvarProcessBar As Object 'local copy
Private mvarStatusLabel As Object 'local copy
Private mvarJoinThisFile As String 'local copy
Private mvarFilesListview As Object 'local copy
Private mvarDestinationPath As String 'local copy
Private mvarFileName As String 'local copy
Private SplittedFilesCount As Integer
Private JoinedFileCount As Integer

Private Const FLAG1 = "."
Private Const FLAG2 = "®"



Public Property Let FileName(ByVal vData As String)
    mvarFileName = GetJoinedName(vData, FLAG1)
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


Public Property Let JoinThisFile(ByVal vData As String)
Attribute JoinThisFile.VB_Description = "The path and file name to be splitted"
    mvarJoinThisFile = vData
    FileName = splitObject.GetFileName(vData)
End Property


Public Property Get JoinThisFile() As String
    JoinThisFile = mvarJoinThisFile
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


Public Property Let ErrNumber(ByVal vData As Integer)
Attribute ErrNumber.VB_Description = "if an error occured this property will contain the error number"
    mvarErrNumber = vData
End Property


Public Property Get ErrNumber() As Integer
    ErrNumber = mvarErrNumber
End Property



Public Property Let ErrMessage(ByVal vData As String)
Attribute ErrMessage.VB_Description = "if an error occured this property will contain the error message"
    mvarErrMessage = vData
End Property


Public Property Get ErrMessage() As String
    ErrMessage = mvarErrMessage
End Property


Public Function LocateSplitedFiles()

    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim SplitIsValid As Boolean
        
    If Not IsFileAValidSplit(JoinThisFile, FLAG1, FLAG2) Then
        ErrNumber = -1
        ErrMessage = "The file you selected is not a splitted file!"
        Exit Function
    Else
        If Val(Mid$(JoinThisFile, InStr(1, JoinThisFile, FLAG2) + 1)) <> 1 Then
            ErrNumber = -1
            ErrMessage = "The file you selected is not the first splitted file!"
            Exit Function
        End If
    End If
    
    SplitIsValid = True
    
    If LCase(Left(JoinThisFile, 2)) <> "a:" Then
        For i = 1 To SplittedFilesCount
            If Dir(GetSplitFile(JoinThisFile) & i) = "" Then
               SplitIsValid = False
               Exit For
            End If
        Next
    
        If Not SplitIsValid Then
           ErrNumber = -1
           ErrMessage = "Missing Split File: " & GetSplitFile(JoinThisFile) & i
           Exit Function
        End If
        If TypeName(FilesListview) = "ListView" Then
            FilesListview.ListItems.Clear
            
            For i = 1 To SplittedFilesCount
                splitObject.AddListViewItem FilesListview, GetSplitFile(JoinThisFile) & i, _
                        FileLen(GetSplitFile(JoinThisFile) & i)
            Next
            
        End If
    End If
    Exit Function
ErrorHandle:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
End Function

Public Function JoinFiles()
    On Error GoTo GetErr
    Dim intjoinfilecount As Integer
    
    If SplittedFilesCount = 0 Then
        LocateSplitedFiles
    End If
    If ErrNumber <> 0 Then
        Exit Function
    End If
    
    Dim i As Integer
    Dim ByteArray() As Byte
     
       
    intWriteHandle = FreeFile
    Open DestinationPath & FileName For Binary Access Write As #intWriteHandle
    
    If TypeName(ProcessBar) = "ProgressBar" Then
       ProcessBar.Visible = True
       ProcessBar.Max = SplittedFilesCount
    End If
    
    Screen.MousePointer = vbHourglass
    For i = 1 To SplittedFilesCount
        DoEvents
        
        If OpenNextSplittedFile() <> 0 Then Exit For
        
        ReDim ByteArray(1 To LOF(intReadHandle))
        
        Get #intReadHandle, , ByteArray
        Put #intWriteHandle, , ByteArray
    Next
    Screen.MousePointer = vbDefault
    Exit Function
GetErr:
    Screen.MousePointer = vbDefault
    ErrNumber = Err.Number
    ErrMessage = Err.Description
End Function


Private Function OpenNextSplittedFile() As Integer
        On Error GoTo Errorhand
        
        Dim strNextSplitedFile As String
     
        Close (intReadHandle)
                
        JoinedFileCount = JoinedFileCount + 1
        
        strNextSplitedFile = GetSplitFile(JoinThisFile) & JoinedFileCount
        
        If LCase(Left(strNextSplitedFile, 2)) = "a:" Then
            If MsgBox("Please insert disk #" & JoinedFileCount & "  of  " & SplittedFilesCount & "  in drive A:.", vbInformation + vbOKCancel) = vbCancel Then
               ErrNumber = -1
               ErrMessage = "Process canceled by the user."
               OpenNextSplittedFile = ErrNumber
               Exit Function
            Else
                If Dir(strNextSplitedFile) = "" Then Err.Raise 52
            End If
        End If
        
        
        intReadHandle = FreeFile
        Open strNextSplitedFile For Binary Access Read As #intReadHandle
        
        If TypeName(ProcessBar) = "ProgressBar" Then
           ProcessBar.Value = JoinedFileCount
        End If
        If TypeName(StatusLabel) = "Label" Then
            If Len("Joining File: " & strNextSplitedFile) > 80 Then
                StatusLabel.Caption = "Joining File: " & "..." & Right$(strNextSplitedFile, 51)
            Else
                StatusLabel.Caption = "Joining File: " & strNextSplitedFile
            End If
        End If
        
        If LCase(Left(strNextSplitedFile, 2)) = "a:" Then
            If TypeName(FilesListview) = "ListView" Then
               splitObject.AddListViewItem FilesListview, strNextSplitedFile, LOF(intReadHandle)
            End If
        End If
        DoEvents
        OpenNextSplittedFile = 0
        Exit Function
Errorhand:
      If Err.Number = 52 Then 'bad file name or number
         If MsgBox("Please insert disk #" & JoinedFileCount & "  of  " & SplittedFilesCount & "  in drive A:.", vbInformation + vbOKCancel) = vbCancel Then
            ErrNumber = -1
            ErrMessage = "Process canceled by the user."
            OpenNextSplittedFile = ErrNumber
            Exit Function
         Else
            Resume
         End If
    Else
        ErrNumber = Err.Number
        ErrMessage = Err.Description
        OpenNextSplittedFile = Err.Number
    End If
End Function



Private Sub Class_Initialize()
    JoinedFileCount = 0
    intReadHandle = 0
    intWriteHandle = 0
    ErrMessage = ""
    ErrNumber = 0
    SplittedFilesCount = 0
    DestinationPath = ""
    FileName = ""
    ConfirmDelete = True
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
        Set FilesListview = Nothing
    End If
    
    Close #intWriteHandle
    Close #intReadHandle
    Set splitObject = Nothing
    If ErrNumber <> 0 Then
        DeletePartialJoin
    End If
End Sub

Private Function DeletePartialJoin()
If DestinationPath & FileName <> "" Then
    If Len(Trim(Dir(DestinationPath & FileName))) > 0 Then
        Kill (DestinationPath & FileName)
    End If
End If
End Function


Public Function DeleteSplittedFiles()
    If ErrNumber = 0 And LCase(Left(JoinThisFile, 2)) <> "a:" Then
       If ConfirmDelete Then
          If MsgBox("All splitted files successfully joined, do you wish to delete them?", vbQuestion + vbOKCancel) = vbCancel Then
             Exit Function
          End If
       End If
       Close (intReadHandle)
       Close (intWriteHandle)
       Dim i As Integer
       For i = 1 To SplittedFilesCount
           If Dir(GetSplitFile(JoinThisFile) & i) <> "" Then
              Kill (GetSplitFile(JoinThisFile) & i)
           End If
       Next
       If TypeName(FilesListview) = "ListView" Then
           FilesListview.ListItems.Clear
       End If
    End If
End Function


Private Function IsFileAValidSplit(strFileName, strFlag1 As String, strFlag2 As String) As Boolean
    Dim i As Integer
    Dim blnValid As Boolean
    Dim strsplitcount As String
    Dim FLAG2 As String
    blnValid = True
    
    For i = Len(Trim(strFileName)) To 1 Step -1
        'check #1
        'we expect digit to be the last char
        If i = Len(Trim(strFileName)) And Not IsNumeric(Mid$(strFileName, i, 1)) Then
           blnValid = False
           Exit For
        End If
        If Not IsNumeric(Mid$(strFileName, i, 1)) Then
           'so, if we are here and it is the first non numeric char then
           'it must be the second flag.
           'we need to check this char against the flag passed to this func
           'if it dosen't match then the user picked up the wrong file.
           If FLAG2 = "" And strFlag2 = Mid$(strFileName, i, 1) Then
                FLAG2 = Mid$(strFileName, i, 1)
                'check #2
                'we expect digit after falg2 has the flag
                If Not IsNumeric(Mid$(strFileName, i - 1, 1)) Then
                   blnValid = False
                   Exit For
                End If
           Else
                'ok, the char dosen't match with the flag, we got a wrong file
                If FLAG2 = "" Then
                    blnValid = False
                    Exit For
                Else
                    'if we are here it means that we got the second flag, now
                    'its time to check the first flag, if it desn't match again
                    'we got the wrong file
                   If strFlag1 <> Mid$(strFileName, i, 1) Then
                      blnValid = False
                      Exit For
                   Else
                        'if the flag matched then we got the right file and lets
                        'get out of here
                        Exit For
                   End If
                End If
           End If
        Else
            'any number of digits in between flag1 and flag2 to means total # of
            'splitted files, so lets get the count of splitted files.
            If FLAG2 <> "" And IsNumeric(Mid$(strFileName, i, 1)) Then
                strsplitcount = Mid$(strFileName, i, 1) & strsplitcount
            End If
        End If
    Next
    If FLAG2 = "" Or strsplitcount = "" Then
       blnValid = False
    Else
        SplittedFilesCount = strsplitcount
    End If
    IsFileAValidSplit = blnValid
End Function

Private Function GetSplitFile(strFile As String) As String
    'This function returns the file name upto the second flag
    'we asume we have the right file, because it is alerady checked.
    
    Dim i As Integer
    For i = Len(Trim(strFile)) To 1 Step -1
        
        If Not IsNumeric(Mid$(strFile, i, 1)) Then 'we are at the second flag
            strFile = Mid$(strFile, 1, i)
            Exit For
        End If
    Next
    GetSplitFile = strFile
End Function

Function GetJoinedName(strSplitFile As String, strFlag1 As String) As String
    Dim i As Integer
    For i = Len(Trim(strSplitFile)) To 1 Step -1
        
        If Mid$(strSplitFile, i, 1) = strFlag1 Then  'we are at the second flag
            strSplitFile = Mid$(strSplitFile, 1, i - 1)
            Exit For
        End If
    Next
    GetJoinedName = strSplitFile
End Function
