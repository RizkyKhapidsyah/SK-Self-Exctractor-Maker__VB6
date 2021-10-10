Attribute VB_Name = "ModWF"
Public FileBuf As Integer

Public Function SelfExtract() As Boolean

    On Error Resume Next
    
    Dim Size As String
    Dim SourceName As String
    Dim FileBinary As String
    
    CurrentPosition = 0
    
    M = 0
    
    Do
        
        M = M + 1
        
        Close #3
        Open App.Path & "\" & App.EXEName & ".exe" For Binary As #3
            
            Seek #3, LOF(3) - (256 * 2) - 5 - 41 - 10 + CurrentPosition
            SourceName = String(40, Chr(0))
            Get #3, , SourceName
            
            SourceName = Replace$(SourceName, vbCr, "")
            
            Seek #3, LOF(3) - (256 * 2) - 5 - 11 + CurrentPosition
            Size = String(10, Chr(0))
            Get #3, , Size
            Size = CCur(Size)
            
            Seek #3, LOF(3) - 51 - Size - (256 * 2) - 5 + CurrentPosition
            FileBinary = String(Size, Chr(0))
            Get #3, , FileBinary
            
        Close #3
        
        Close #4
        
        Open FrmMain.ArchiveFName.Text For Binary Access Write As #4
            Put #4, , FileBinary
        Close #4
        
        CurrentPosition = CurrentPosition - Size - 50
        
    Loop Until M >= FileBuf
    
    SelfExtract = True
    
    Exit Function
    
FinaliseError:
    
    MsgBox "An error occured. Header may be damaged. This file could not open.", vbCritical, "Error"
    
    SelfExtract = False

End Function
