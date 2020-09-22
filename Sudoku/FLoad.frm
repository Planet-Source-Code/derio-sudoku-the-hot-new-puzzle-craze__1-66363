VERSION 5.00
Begin VB.Form FLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load File ..."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox lstFile 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "FLoad.frx":0000
      Left            =   780
      List            =   "FLoad.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FLoad.frx":0004
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Tag = ""
  Hide
End Sub

Private Sub cmdOK_Click()
  If Me.lstFile.ListIndex <> -1 Then
    Tag = Me.lstFile.ItemData(Me.lstFile.ListIndex)
  End If
  Hide
End Sub

Private Sub lstFile_DblClick()
  If Me.lstFile.ListIndex <> -1 Then
    Tag = Me.lstFile.ItemData(Me.lstFile.ListIndex)
    Hide
  End If
End Sub
