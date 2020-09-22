VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ..."
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2025
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright (C) Derio 2006"
      Height          =   255
      Left            =   885
      TabIndex        =   3
      Top             =   630
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   870
      TabIndex        =   2
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   675
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FAbout.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Hide
End Sub
