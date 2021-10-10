VERSION 5.00
Begin VB.Form FrmDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Directory"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "FrmDirectory.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox SetDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.DirListBox SetDir 
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton OKCmd 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton CancelCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastDrive As String

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub OKCmd_Click()
    FrmMain.ExtractPath.Text = SetDir.Path
    Unload Me
End Sub

Private Sub SetDir_Change()
    SetFiles = SetDir
End Sub

Private Sub SetDrive_Change()
    
    On Error GoTo FinaliseError
    
    LastDrive = SetDrive
    SetDir = SetDrive
    
    Exit Sub
    
FinaliseError:
    
    MsgBox SetDrive.List(SetDrive.ListIndex) & vbNewLine & vbNewLine & "The device is not ready.", vbCritical, SetDrive.List(SetDrive.ListIndex)
    SetDrive = LastDrive
    
End Sub

