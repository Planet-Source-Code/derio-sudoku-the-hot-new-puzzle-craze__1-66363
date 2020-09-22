VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sudoku"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6360
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FMain.frx":030A
   ScaleHeight     =   6360
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6600
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   81
      Top             =   4200
      Width           =   600
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "9"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   90
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "8"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   89
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "7"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   88
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   87
         Top             =   180
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   86
         Top             =   180
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   85
         Top             =   180
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   84
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   83
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   80
      Left            =   5640
      TabIndex        =   80
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   79
      Left            =   4980
      TabIndex        =   79
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   78
      Left            =   4320
      TabIndex        =   78
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   77
      Left            =   3540
      TabIndex        =   77
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   76
      Left            =   2880
      TabIndex        =   76
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   75
      Left            =   2220
      TabIndex        =   75
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   74
      Left            =   1440
      TabIndex        =   74
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   73
      Left            =   780
      TabIndex        =   73
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   72
      Left            =   120
      TabIndex        =   72
      Top             =   5640
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   71
      Left            =   5640
      TabIndex        =   71
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   70
      Left            =   4980
      TabIndex        =   70
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   69
      Left            =   4320
      TabIndex        =   69
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   68
      Left            =   3540
      TabIndex        =   68
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   67
      Left            =   2880
      TabIndex        =   67
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   66
      Left            =   2220
      TabIndex        =   66
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   65
      Left            =   1440
      TabIndex        =   65
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   64
      Left            =   780
      TabIndex        =   64
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   63
      Left            =   120
      TabIndex        =   63
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   62
      Left            =   5640
      TabIndex        =   62
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   61
      Left            =   4980
      TabIndex        =   61
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   60
      Left            =   4320
      TabIndex        =   60
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   59
      Left            =   3540
      TabIndex        =   59
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   58
      Left            =   2880
      TabIndex        =   58
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   57
      Left            =   2220
      TabIndex        =   57
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   56
      Left            =   1440
      TabIndex        =   56
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   55
      Left            =   780
      TabIndex        =   55
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   54
      Left            =   120
      TabIndex        =   54
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   53
      Left            =   5640
      TabIndex        =   53
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   52
      Left            =   4980
      TabIndex        =   52
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   51
      Left            =   4320
      TabIndex        =   51
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   50
      Left            =   3540
      TabIndex        =   50
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   49
      Left            =   2880
      TabIndex        =   49
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   48
      Left            =   2220
      TabIndex        =   48
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   47
      Left            =   1440
      TabIndex        =   47
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   46
      Left            =   780
      TabIndex        =   46
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   45
      Left            =   120
      TabIndex        =   45
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   44
      Left            =   5640
      TabIndex        =   44
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   43
      Left            =   4980
      TabIndex        =   43
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   42
      Left            =   4320
      TabIndex        =   42
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   41
      Left            =   3540
      TabIndex        =   41
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   40
      Left            =   2880
      TabIndex        =   40
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   39
      Left            =   2220
      TabIndex        =   39
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   38
      Left            =   1440
      TabIndex        =   38
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   37
      Left            =   780
      TabIndex        =   37
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   36
      Left            =   120
      TabIndex        =   36
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   35
      Left            =   5640
      TabIndex        =   35
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   34
      Left            =   4980
      TabIndex        =   34
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   33
      Left            =   4320
      TabIndex        =   33
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   32
      Left            =   3540
      TabIndex        =   32
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   31
      Left            =   2880
      TabIndex        =   31
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   30
      Left            =   2220
      TabIndex        =   30
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   29
      Left            =   1440
      TabIndex        =   29
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   28
      Left            =   780
      TabIndex        =   28
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   27
      Left            =   120
      TabIndex        =   27
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   26
      Left            =   5640
      TabIndex        =   26
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   25
      Left            =   4980
      TabIndex        =   25
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   24
      Left            =   4320
      TabIndex        =   24
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   23
      Left            =   3540
      TabIndex        =   23
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   22
      Left            =   2880
      TabIndex        =   22
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   21
      Left            =   2220
      TabIndex        =   21
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   20
      Left            =   1440
      TabIndex        =   20
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   19
      Left            =   780
      TabIndex        =   19
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   17
      Left            =   5640
      TabIndex        =   17
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   16
      Left            =   4980
      TabIndex        =   16
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   15
      Left            =   4320
      TabIndex        =   15
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   14
      Left            =   3540
      TabIndex        =   14
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   13
      Left            =   2880
      TabIndex        =   13
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   12
      Left            =   2220
      TabIndex        =   12
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   11
      Left            =   1440
      TabIndex        =   11
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   10
      Left            =   780
      TabIndex        =   10
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   4980
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   3540
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   2220
      TabIndex        =   3
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Library ..."
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save last position"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore saved position"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "Advanced"
      Begin VB.Menu mnuSaveLibrary 
         Caption         =   "Save Last Position as Library ..."
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavigation 
         Caption         =   "Navigation On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuResolvePuzzle 
         Caption         =   "Resolve Puzzle"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRule 
         Caption         =   "The Rule"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************
'* Title  : Sudoku 1.0       *
'* Author : Derio            *
'* Type   : Puzzle Game      *
'* Stamp  : 22 Aug 2006      *
'*****************************

Private Const SavingFileName = "SUDOKU.SAV"
Private Const LibraryExt = "SLB"

Private CurrentIndex As Integer
Private CurrentSelection As Integer

Private MoveHistory As Collection

Private Sub Form_Load()
'** Start the game

  InitNewGame
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If MsgBox("Are you sure to exit?", vbQuestion + vbYesNo) = vbNo Then
    Cancel = True
    Exit Sub
  End If
End Sub

Private Sub lblOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'** Select option value

Dim I As Integer
Dim Complete As Boolean

  With lblSudoku(CurrentIndex)
    If Button = vbLeftButton Then
      .Caption = lblOption(Index).Caption
      .BackStyle = 1
      If CheckComplete() Then Complete = True
    Else
      .Caption = ""
      .BackStyle = 0
    End If
  End With
  
  pctOption.Visible = False
  If Complete Then
    MsgBox "Congratulation, you just solve the puzzle!", vbInformation
    
  Else
    CheckSudoku
  End If
End Sub

Private Sub lblSudoku_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'** Main function of Sudoku

Dim I As Integer
Dim ArrSudoku(1 To 9) As Integer

  If Shift = 1 Then '* protect Sudoku item
    With lblSudoku(Index)
      If .Caption <> "" Then
        If .Tag = "" Then
          .Tag = "1"
          .BackStyle = 0
        Else
          .Tag = ""
          .BackStyle = 1
        End If
      End If
    End With
  
  Else
    '*if the item protected, you can not change the value
    If lblSudoku(Index).Tag = "1" Then Exit Sub
    
    '* left click to change value, right click to lear
    If Button = vbLeftButton Then
      CurrentIndex = Index
      
      With pctOption
        CurrentSelection = 0
        .Left = lblSudoku(Index).Left
        .Top = lblSudoku(Index).Top
        
        GetSudokuItemList Index, ArrSudoku()
        
        '* enable all of the options
        For I = 1 To 9
          With lblOption(I - 1)
            If .BackColor = vbBlack Then
              .BackColor = .ForeColor
              .ForeColor = vbBlack
            End If
            .Enabled = (ArrSudoku(I) = 1)
            .FontBold = .Enabled
          End With
        Next I
        
        If lblSudoku(Index).Caption <> "" Then
          With lblOption(lblSudoku(Index).Caption - 1)
            .Enabled = True
            .ForeColor = .BackColor
            .BackColor = vbBlack
            .FontBold = True
          End With
        End If
      
        If Not .Visible Then .Visible = True
      End With
    
    Else 'clear value
      With lblSudoku(Index)
        .Caption = ""
        .BackStyle = 0
      End With
      
      CheckSudoku
      
      If Me.pctOption.Visible Then Me.pctOption.Visible = False
    End If
  End If
End Sub

Private Sub mnuAbout_Click()
'** Show something about me :-)

Dim fTemp As FAbout

  Set fTemp = New FAbout
  With fTemp
    fTemp.Show vbModal
  End With
  
  Unload fTemp
  Set fTemp = Nothing
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuLoad_Click()
'** Load legendary Sudoku puzzle

Dim fTemp As FLoad
Dim hFile As Integer
Dim strFileName As String
Dim strTemp As String
Dim arrFileName() As String
Dim I As Integer

  If MsgBox("Are you sure to load a new libaray " & _
            "and cancel current Sudoku Puzzle?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            
  Set fTemp = New FLoad
  With fTemp
    '* get list of library
    ReDim arrFileName(0)
    I = 0
    strFileName = Dir(App.Path & "\*." & LibraryExt)
    While strFileName <> ""
      I = I + 1
      ReDim Preserve arrFileName(I)
      arrFileName(I) = strFileName
      hFile = FreeFile
      Open App.Path & "\" & strFileName For Input As #hFile
      Line Input #hFile, strTemp
      .lstFile.AddItem strTemp
      .lstFile.ItemData(.lstFile.NewIndex) = I
      Close #hFile
      strFileName = Dir()
    Wend
    
    .Show vbModal
    If .Tag <> "" Then
      ClearSudokuBoard
      
      '* get the initial state
      strFileName = arrFileName(.Tag)
      hFile = FreeFile
      Open App.Path & "\" & strFileName For Input As #hFile
      Line Input #hFile, strTemp
      Caption = "Sudoku - " & strTemp
      While Not EOF(hFile)
        Input #hFile, strTemp
        I = Val(Left(strTemp, 2))
        With Me.lblSudoku(I - 1)
          .Caption = Trim(Mid(strTemp, 4))
          .Tag = "1"
          .BackStyle = 0
        End With
      Wend
    End If
  End With
  
  Unload fTemp
  Set fTemp = Nothing
End Sub

Private Sub mnuResolvePuzzle_Click()
Dim Finished As Boolean
Dim StepCount As Long

  If MsgBox("Are you sure to let me solve this puzzle?", _
            vbQuestion + vbYesNo) = vbNo Then Exit Sub
  
  Finished = False
  StepCount = 0
  ResolveUsingLinear 0, Finished, StepCount
  If Finished Then
    MsgBox "Puzzle resolved with " & StepCount & " steps!", vbInformation
    
  Else
    MsgBox "I'm sorry, I can't resove this puzzle", vbInformation
  End If
End Sub

Private Sub mnuRestore_Click()
'** Load last position

Dim hFile As Integer
Dim I As Integer
Dim strTemp As String

  '* check existing saved position
  If Dir(App.Path & "\" & SavingFileName) = "" Then
    MsgBox "Saving position not found ...", vbInformation
    Exit Sub
  End If
  
  ClearSudokuBoard
  
  '* restore saved position
  hFile = FreeFile
  Open App.Path & "\" & SavingFileName For Input As #hFile
  While Not EOF(hFile)
    Input #hFile, strTemp
    I = Left(strTemp, 2)
    If IsNumeric(I) Then
      With Me.lblSudoku(I - 1)
        .Caption = Trim(Mid(strTemp, 4, 1))
        .Tag = Trim(Mid(strTemp, 6, 1))
        If .Caption <> "" And .Tag = "" Then .BackStyle = 1
      End With
    End If
  Wend
  Close #hFile
  
  MsgBox "Last position restored ...", vbInformation
End Sub

Private Sub mnuNavigation_Click()
'** Activating or Deactivating navigation key (for help ...)

  If Not mnuNavigation.Checked Then
    mnuNavigation.Checked = True
  Else
    mnuNavigation.Checked = False
  End If
End Sub

Private Sub mnuSave_Click()
'** Save current position

Dim hFile As Integer
Dim I As Integer

  '* check last position
  If Dir(App.Path & "\" & SavingFileName) <> "" Then
    Kill App.Path & "\" & SavingFileName
  End If
  
  '* saved last position
  hFile = FreeFile
  Open App.Path & "\" & SavingFileName For Output As #hFile
  For I = 1 To Me.lblSudoku.Count
    If Me.lblSudoku(I - 1).Caption <> "" Then
      Print #hFile, Format(I, "00") & " " & _
                    Me.lblSudoku(I - 1).Caption & " " & _
                    Me.lblSudoku(I - 1).Tag
    End If
  Next I
  Close #hFile
  
  MsgBox "Last Position Saved ...", vbInformation
End Sub

Private Sub mnuSaveLibrary_Click()
'** Saving Sudoku as library

Dim hFile As Integer
Dim I As Integer
Dim strCaption As String
Dim MaxIndex As Integer
Dim strFileName As String
Dim strTemp As String

  '* get the title
  strCaption = InputBox("Caption", "Sudoku Library Title")
  If strCaption = "" Then
    MsgBox "Saving Library Cancelled!" & vbCrLf & vbCrLf & _
           "Please input some thing to be Sudoku Library Title", _
           vbCritical
    Exit Sub
  End If
  
  '* get the maximum id
  MaxIndex = 0
  strFileName = Dir(App.Path & "\SUDOKU-*." & LibraryExt)
  While strFileName <> ""
    strTemp = Mid(strFileName, 8, 3)
    If MaxIndex < Val(strTemp) Then
      MaxIndex = Val(strTemp)
    End If
    strFileName = Dir()
  Wend
  
  Do
    MaxIndex = MaxIndex + 1
    strFileName = "SUDOKU-" & Format(MaxIndex, "000") & "." & LibraryExt
  
    '* check last position
  Loop Until Dir(App.Path & "\" & strFileName) = ""
  
  '* saved last position as library
  hFile = FreeFile
  Open App.Path & "\" & strFileName For Output As #hFile
  Print #hFile, strCaption
  For I = 1 To Me.lblSudoku.Count
    If Me.lblSudoku(I - 1).Caption <> "" Then
      Print #hFile, Format(I, "00") & " " & _
                    Me.lblSudoku(I - 1).Caption & " " & _
                    Me.lblSudoku(I - 1).Tag
    End If
  Next I
  Close #hFile
  
  MsgBox "Library saved ...", vbInformation
End Sub






Private Sub ClearSudokuBoard()
'** Clear all of the sudoku item
Dim I As Integer

  For I = 1 To Me.lblSudoku.Count
    With Me.lblSudoku(I - 1)
      .Caption = ""
      .Tag = ""
      If Not .Visible Then .Visible = True
      .BackStyle = 0
    End With
  Next I
  If Me.pctOption.Visible Then Me.pctOption.Visible = False
End Sub

Private Function GetBox(ByVal Index As Integer) As Integer
'** Get the box index

  GetBox = ((GetRow(Index) - 1) \ 3) * 3 + ((GetColumn(Index) - 1) \ 3) + 1
End Function

Private Function GetRow(ByVal Index As Integer) As Integer
'** Get the row index

  GetRow = (Index \ 9) + 1
End Function

Private Function GetColumn(ByVal Index As Integer) As Integer
'** Get the column pos

  GetColumn = (Index Mod 9) + 1
End Function

Private Function GetIndex(ByVal X As Integer, ByVal Y As Integer) As Integer
'** Get the index base on X and Y

  GetIndex = (Y - 1) * 9 + X - 1
End Function

Private Function GetBoxIndex(ByVal BoxIndex As Integer, ByVal Index As Integer) As Integer
'** Get the index base on BoxIndex and Index

  GetBoxIndex = Index + _
                ((Index - 1) \ 3) * 6 + _
                (BoxIndex - 1) * 3 + _
                ((BoxIndex - 1) \ 3) * 18 - 1
End Function

Private Sub GetSudokuItemList(ByVal Index As Integer, ArraySudoku() As Integer)
'** Get the possible Sudoku item

Dim I As Integer
Dim J As Integer
Dim strTemp As String

  '* init array sudoku
  For I = 1 To 9
    ArraySudoku(I) = 1
  Next I
  
  '* disable the option if the same one at the same column has choosen
  J = GetColumn(Index)
  For I = 1 To 9
    strTemp = lblSudoku(GetIndex(J, I)).Caption
    If strTemp <> "" Then
      ArraySudoku(Val(strTemp)) = 0
    End If
  Next I
  
  '* disable the option if the same one at the same row has choosen
  J = GetRow(Index)
  For I = 1 To 9
    strTemp = lblSudoku(GetIndex(I, J)).Caption
    If strTemp <> "" Then
      ArraySudoku(Val(strTemp)) = 0
    End If
  Next I
  
  '* disable the option if the same one at the same box has choosen
  J = GetBox(Index)
  For I = 1 To 9
    strTemp = lblSudoku(GetBoxIndex(J, I)).Caption
    If strTemp <> "" Then
      ArraySudoku(Val(strTemp)) = 0
    End If
  Next I
  
End Sub

Private Function CheckSudoku() As Boolean
'** Check sudoku item

Dim I As Integer

  CheckSudoku = True
  If Not Me.mnuNavigation.Checked Then Exit Function
  
  For I = 1 To Me.lblSudoku.Count
    If Me.lblSudoku(I - 1).Caption = "" Then
      If Not SudokuOK(I - 1) Then
        Me.lblSudoku(I - 1).Visible = False
        CheckSudoku = False
      Else
        Me.lblSudoku(I - 1).Visible = True
      End If
    End If
  Next I
End Function

Private Function SudokuOK(ByVal Index As Integer) As Boolean
'** Check if Sudoku item has choice

Dim I As Integer
Dim J As Integer
Dim arrTemp(1 To 9) As Integer
Dim strTemp As String

  SudokuOK = False
  '* enable all of the options
  For I = 1 To 9
    arrTemp(I) = 1
  Next I
  
  GetSudokuItemList Index, arrTemp()
  
  For I = 1 To 9
    If arrTemp(I) <> 0 Then
      SudokuOK = True
      Exit Function
    End If
  Next I
End Function

Private Sub InitNewGame()
'** Init new game
  
  ClearSudokuBoard
  Set MoveHistory = Nothing
  Set MoveHistory = New Collection
End Sub

Private Function CheckComplete() As Boolean
'** Check for the completeness

Dim I As Integer

  For I = 1 To Me.lblSudoku.Count
    If Me.lblSudoku(I - 1).Caption = "" Then
      CheckComplete = False
      Exit Function
    End If
  Next I
  
  CheckComplete = True
End Function


Private Sub ResolveUsingLinear(ByVal Index As Integer, Finished As Boolean, StepCount As Long)
'** Solving puzzle using linear method

Dim I As Integer
Dim arrTemp(1 To 9) As Integer

  StepCount = StepCount + 1
  If Index > 80 Then
    Finished = True
    
  ElseIf Me.lblSudoku(Index).Caption = "" Then
    GetSudokuItemList Index, arrTemp()
    For I = 1 To 9
      If arrTemp(I) = 1 Then
        Me.lblSudoku(Index).Caption = I
        Me.lblSudoku(Index).BackStyle = 1
        DoEvents
        If CheckSudoku() Then
          ResolveUsingLinear Index + 1, Finished, StepCount
          If Not Finished Then
            Me.lblSudoku(Index).Caption = ""
            Me.lblSudoku(Index).BackStyle = 0
            CheckSudoku
            DoEvents
            
          Else
            Exit Sub
          End If
          
        Else
          Me.lblSudoku(Index).Caption = ""
          Me.lblSudoku(Index).BackStyle = 0
          CheckSudoku
          DoEvents
        End If
      End If
    Next I
    
  ElseIf Not Finished Then
    ResolveUsingLinear Index + 1, Finished, StepCount
  End If
End Sub
