Attribute VB_Name = "ModSE"
Public FileBuf As Integer

Sub SelfExtract()

    On Error Resume Next
    
    Dim Size As String
    Dim SourceName As String
    Dim FileBinary As String
    
    CurrentPosition = 0
    
    M = 0
    
    Do
        
        M = M + 1
        
'    If Command <> "" Then
'        Open Command For Binary As #3
'    Else
        Open App.Path & "\" & App.EXEName & ".exe" For Binary As #3
'    End If
            
            Seek #3, LOF(3) - (256 * 2) - 5 - 41 - 10 + CurrentPosition
            SourceName = String(40, Chr(0))
            Get #3, , SourceName
            
            SourceName = Replace$(SourceName, vbCr, "")
            frmSelfExtract.lblFiles.Caption = "Extracting " & SourceName & "..."
            frmSelfExtract.lblFiles.Refresh
            
            Seek #3, LOF(3) - (256 * 2) - 5 - 11 + CurrentPosition
            Size = String(10, Chr(0))
            Get #3, , Size
            Size = CCur(Size)
            
            Seek #3, LOF(3) - 51 - Size - (256 * 2) - 5 + CurrentPosition
            FileBinary = String(Size, Chr(0))
            Get #3, , FileBinary
            
        Close #3
        
        Open FrmMain.ExtractPath.Text & "\" & SourceName For Binary Access Write As #4
            Put #4, , FileBinary
        Close #4
        
        CurrentPosition = CurrentPosition - Size - 50
        
    Loop Until M >= FileBuf
    
    End
    
    Exit Sub
    
FinaliseError:
    
    Result = MsgBox("An error occured. Header may be damaged." & vbCrLf & "Do you want to abort/retry?", vbAbortRetryIgnore + vbExclamation, "Error")
    
    If Result = vbRetry Then
        Resume
    ElseIf Result = vbIgnore Then
        Resume Next
    ElseIf Result = vbAbort Then
        End
    End If

End Sub
