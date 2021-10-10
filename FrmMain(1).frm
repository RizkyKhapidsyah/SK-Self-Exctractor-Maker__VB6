VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-Extractor"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain(1).frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAbout 
      Caption         =   "&About"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse..."
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
      Left            =   2700
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox ExtractPath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Files"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract"
      Default         =   -1  'True
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
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmMain(1).frx":0CCA
      Height          =   630
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   4770
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extract to:"
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
      Left            =   150
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3675
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Please specify a extraction location, and the click on the Extract button."
      Height          =   450
      Left            =   150
      TabIndex        =   5
      Top             =   240
      Width           =   3720
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


Private Sub CmdAbout_Click()
    MsgBox "Self-Extractor, Created by Pre-Instinct® Software" & vbNewLine & "Produced by Mark Withers. Contact him on his email at:" & vbNewLine & vbNewLine & "NeoBPI@Yahoo.com" & vbNewLine & vbNewLine & "Or visit the website at:" & vbNewLine & vbNewLine & "www.Pre-Instinct-Software.20m.com (For more details.)"
End Sub

Private Sub CmdBrowse_Click()
    On Error Resume Next
    FrmDirectory.Show 1
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Function CheckDir(Directory As String) As Boolean
    
    On Error GoTo FinaliseError
    ChDir Directory
    CheckDir = True
    
    Exit Function
    
FinaliseError:
    
    CheckDir = False

End Function

Private Sub cmdExtract_Click()
    
    If CheckDir(ExtractPath.Text) = False Then MsgBox "Invalid folder name. Please specify a valid location.", vbCritical, "Self-Extractor": Exit Sub
    
    cmdView.Enabled = False
    cmdExtract.Enabled = False
    SelfExtract
    
End Sub

Private Sub cmdView_Click()
    On Error Resume Next
    frmFiles.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Form_Load()

'    If Command <> "" Then
'        Me.Caption = "Self-Extractor - " & OnlyFileName(Command)
'    Else
        Me.Caption = "Self-Extractor - " & App.EXEName & ".exe"
'    End If

    On Error GoTo FinaliseError
    
    Dim SourceFile As String
    Dim SourceName As String
    Dim Size As String
    
    CurrentPosition = 0
    M = 0
    
    Close #1
'    If Command <> "" Then
'        Open Command For Binary As #1
'    Else
        Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
'    End If
        
            Seek #1, LOF(1) - 6 - (256 * 2)
            SourceFile = String(5, Chr(0))
            Get #1, , SourceFile
        
            SourceFile = Replace$(SourceFile, vbCr, "")
            FileBuf = Val(SourceFile)
            lblFiles.Caption = "This archive contains " & FileBuf & " file(s)."
            
        Close #1
    
    Do
    
        M = M + 1
        
'    If Command <> "" Then
'        Open Command For Binary As #1
'    Else
        Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
'    End If
    
            Seek #1, LOF(1) - (256 * 2) - 5 - 41 - 10 + CurrentPosition
            SourceName = String(40, Chr(0))
            Get #1, , SourceName
            
            Seek #1, LOF(1) - (256 * 2) - 5 - 11 + CurrentPosition
            Size = String(10, Chr(0))
            Get #1, , Size
            Size = CCur(Size)
            
        Close #1
            
        SourceName = Replace$(SourceName, vbCr, "")
        frmFiles.lstFiles.AddItem SourceName
        
        CurrentPosition = CurrentPosition - Size - 50
    
    Loop Until M >= FileBuf
    
    Exit Sub
FinaliseError:
    MsgBox "This file is damaged or it doesn't include any files.", vbCritical, "Error"
    End
End Sub
