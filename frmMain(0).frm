VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Self-Extractor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmMain(0).frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton AddCmd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ArchiveFName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click on browse to specify a new archive:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function OnlyFileName(file As String) As String
    
    Dim CheckPos As Long
    Dim FNameLen As Long
    
    If InStr(file, "\") = 0 And InStr(file, "/") = 0 Then OnlyFileName = file: Exit Function
    
    CheckPos = 1
    
    Do
        
        FNameLen = CheckPos
        
        If InStr(CheckPos + 1, file, "\") = 0 Then
            CheckPos = InStr(CheckPos + 1, file, "/")
        Else
            CheckPos = InStr(CheckPos + 1, file, "\")
        End If
    
    Loop Until CheckPos = 0
    
    OnlyFileName = Right(file, Len(file) - Len(Left(file, FNameLen)))
    
End Function

Private Sub CmdBrowse_Click()
    
    On Error GoTo FinaliseError
    
    Dlg.CancelError = True
    Dlg.Filter = "EXE Archives|*.exe|"
    Dlg.Flags = cdlOFNFileMustExist
    Dlg.ShowSave
    
    If Dlg.FileName = "" Then Exit Sub
    
    ArchiveFName = Dlg.FileName

FinaliseError:

End Sub

Private Sub AddCmd_Click()
    
    On Error GoTo FinaliseError
    
    Dlg.CancelError = True
    Dlg.Filter = "All Files|*.*|"
    Dlg.Flags = cdlOFNFileMustExist
    Dlg.ShowOpen
    
    If Dlg.FileName = "" Then Exit Sub
    
    For i = 0 To lstFiles.ListCount - 1
        If LCase$(OnlyFileName(Dlg.FileName)) = LCase$(OnlyFileName(lstFiles.List(i))) Then MsgBox "A file with the same name exists!", vbExclamation, "Error": Exit Sub
    Next i
    
    lstFiles.AddItem Dlg.FileName

FinaliseError:

End Sub

Private Sub CmdOK_Click()
    
    If ArchiveFName.Text = "" Then MsgBox "Please specify a new EXE - Archive.", vbExclamation, "Self-Extractor": ArchiveFName.SetFocus: Exit Sub
    If lstFiles.ListCount = 0 Then MsgBox "Please specify at least one file to add to the archive.", vbExclamation, "Self-Extractor": Exit Sub
    
    'If FileExist(ArchiveFName.Text) = True Then If KillFile(ArchiveFName.Text) = False Then MsgBox "Error, could not completly over-right file. Of a error of this, the new archive specified could see a change in size.", vbCritical, "Error"
    
    If SelfExtract = True Then
        If AddToSelfExtract(ArchiveFName.Text, Me.lstFiles, ArchiveFName.Text) = True Then
            MsgBox "Your new self-extracting archive has now been created.", vbInformation, "Self-Extractor"
        End If
    End If
    
End Sub

Private Sub CmdClose_Click()
    End
End Sub

Private Sub CmdRemove_Click()
    On Error Resume Next
    lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub Form_Load()
    If Command <> "" Then ArchiveFName.Text = Command
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
