VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Search"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000003&
         Caption         =   "All"
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000003&
         Caption         =   "Individual"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000003&
         Caption         =   "By Period"
         Height          =   735
         Left            =   4080
         TabIndex        =   3
         Top             =   720
         Width           =   4935
         Begin VB.PictureBox DTPicker2 
            Height          =   375
            Left            =   2880
            ScaleHeight     =   315
            ScaleWidth      =   1875
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.PictureBox DTPicker1 
            Height          =   375
            Left            =   240
            ScaleHeight     =   315
            ScaleWidth      =   1995
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000003&
         Caption         =   "By Name "
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3735
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   3255
         End
      End
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   1800
      Width           =   9135
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call fConn
SQL "Select * from tblemployee"
End Sub
