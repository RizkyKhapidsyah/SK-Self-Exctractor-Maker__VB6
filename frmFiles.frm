VERSION 5.00
Begin VB.Form frmFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listed files in archive"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   FillColor       =   &H8000000F&
   Icon            =   "frmFiles.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUnload_Click()
    Hide
End Sub
