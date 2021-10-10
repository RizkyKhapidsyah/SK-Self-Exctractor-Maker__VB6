Attribute VB_Name = "ModMK"
Public Function FileExist(FileName As String) As Boolean
    
    On Error GoTo FinaliseError
    
    If FileLen(FileName) <> 0 Then
        FileExist = True
    Else
        FileExist = False
    End If
    
FinaliseError:
    
    FileExist = False
    
End Function

Public Function KillFile(FileName As String) As Boolean
    
    On Error GoTo FinaliseError
    
    Kill FileName
    KillFile = True
    
    Exit Function
    
FinaliseError:

    KillFile = False

End Function

Function AddToSelfExtract(SelfExtract As String, WhatFile As ListBox, SaveAs As String) As Boolean

    On Error GoTo FinaliseError

    Dim SourceBinary As String
    Dim FileBinary As String
    Dim SourceFName As String
    Dim LabeledText As String
    Dim ArchiveAbout As String
    Dim SourceName As String

    Close #1
    Open SelfExtract For Binary As #1

        FileBinary = String(LOF(1), Chr(0))

        Get #1, , FileBinary

    Close #1

    Close #1
    Open SaveAs For Output As #1

        wholePrint = FileBinary

        For M = 0 To WhatFile.ListCount - 1

            SourceName = FrmMain.OnlyFileName(WhatFile.List(M))

            FrmMain.Caption = "Reading " & FrmMain.OnlyFileName(WhatFile.List(M)) & "..."
            FrmMain.Refresh

            Close #2
            Open WhatFile.List(M) For Binary As #2

                SourceBinary = String(LOF(2), Chr(0))
                Get #2, , SourceBinary

                Size = LOF(2)
                SourceName = String(40 - Len(SourceName), vbCr) & SourceName
                Size = String(10 - Len(Size), "0") & Size
                wholePrint = wholePrint & SourceBinary & SourceName & Size

            Close #2

        Next M

        LabeledText = "Pre-Instinct® Software Self Extractor"
        LabeledText = String(256 - Len(LabeledText), vbTab) & LabeledText

        ArchiveAbout = "Created by Pre-Instinct® Software and may not be copied for any reason without consent of the producer of this product."
        ArchiveAbout = String(256 - Len(ArchiveAbout), vbTab) & ArchiveAbout

        SourceFName = WhatFile.ListCount
        SourceFName = String(5 - Len(SourceFName), vbCr) & SourceFName

        Print #1, wholePrint & SourceFName & LabeledText & ArchiveAbout

    Close #1

    FrmMain.Caption = "Self-Extractor"
    FrmMain.Refresh

    AddToSelfExtract = True

    Exit Function

FinaliseError:

    MsgBox "An error occured while creating self extractor.", vbCritical, "Error"
    AddToSelfExtract = False

End Function
