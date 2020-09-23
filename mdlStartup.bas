Attribute VB_Name = "mdlStartUp"
Public gDestination As String

'this function extracts file path from a path with filename
Function GetFilePath(strFileWithPath As String) As String

    Dim strPath As String
    Do
        strPath = strPath & Mid$(strFileWithPath, 1, InStr(1, strFileWithPath, "\"))
        strFileWithPath = Mid$(strFileWithPath, _
                          IIf(InStr(1, strFileWithPath, "\") > 0, _
                              InStr(1, strFileWithPath, "\") + 1, 1))
        
    Loop While InStr(1, strFileWithPath, "\") <> 0
    GetFilePath = strPath
End Function


'this function gets value from windows registry
Function GetRegistryValue(strAppName As String, strSection As String, strKey As String, Default As Variant) As String


    On Error GoTo ErrorHandle
    'Get the value from registry
    GetRegistryValue = GetSetting(strAppName, strSection, strKey, Default)
    Exit Function
ErrorHandle:
        GetRegistryValue = Default
        MsgBox Err.Description
End Function

'this function saves value to windows registry
Function SaveRegistryValue(strAppName As String, strSection As String, strKey As String, strNewValue As String) As Boolean
'Updates values in the registry

    On Error GoTo ErrorHandle
    
    SaveSetting strAppName, strSection, strKey, strNewValue
      
    SaveRegistryValue = True
      
    Exit Function
ErrorHandle:
        PutRegistryValue = False
        MsgBox Err.Description
End Function

