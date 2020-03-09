VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransact 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransact.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   6735
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   7095
      Begin VB.Frame Framex 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         Height          =   6375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6855
         Begin TabDlg.SSTab SSTab1 
            Height          =   6015
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   10610
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            BackColor       =   -2147483645
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Task   "
            TabPicture(0)   =   "frmTransact.frx":400BC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label12"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label11"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label4"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label3"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Frame3"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Frame2"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "Deduction"
            TabPicture(1)   =   "frmTransact.frx":400D8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cboDeductionType"
            Tab(1).Control(1)=   "Text5"
            Tab(1).Control(2)=   "Command2"
            Tab(1).Control(3)=   "Command3"
            Tab(1).Control(4)=   "DTPicker2"
            Tab(1).Control(5)=   "DTPicker1"
            Tab(1).Control(6)=   "ListView3"
            Tab(1).Control(7)=   "Line4"
            Tab(1).Control(8)=   "Label40"
            Tab(1).Control(9)=   "Label39"
            Tab(1).Control(10)=   "Label38"
            Tab(1).Control(11)=   "Label5"
            Tab(1).Control(12)=   "Label6"
            Tab(1).Control(13)=   "Label7"
            Tab(1).Control(14)=   "Label8"
            Tab(1).ControlCount=   15
            TabCaption(2)   =   "Allowance"
            TabPicture(2)   =   "frmTransact.frx":400F4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame4"
            Tab(2).Control(1)=   "Command5"
            Tab(2).Control(2)=   "Command4"
            Tab(2).Control(3)=   "ListView4"
            Tab(2).Control(4)=   "Line5"
            Tab(2).Control(5)=   "Label43"
            Tab(2).Control(6)=   "Label42"
            Tab(2).Control(7)=   "Label13"
            Tab(2).Control(8)=   "Label14"
            Tab(2).Control(9)=   "Label17"
            Tab(2).Control(10)=   "Label18"
            Tab(2).Control(11)=   "Label19"
            Tab(2).Control(12)=   "Label20"
            Tab(2).ControlCount=   13
            TabCaption(3)   =   "Computation"
            TabPicture(3)   =   "frmTransact.frx":40110
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Command1"
            Tab(3).Control(1)=   "Frame5"
            Tab(3).Control(2)=   "Label37"
            Tab(3).Control(3)=   "Label36"
            Tab(3).Control(4)=   "Label35"
            Tab(3).Control(5)=   "Label34"
            Tab(3).Control(6)=   "Label33"
            Tab(3).Control(7)=   "Label32"
            Tab(3).Control(8)=   "Label31"
            Tab(3).Control(9)=   "Label30"
            Tab(3).Control(10)=   "Line3"
            Tab(3).Control(11)=   "Line2"
            Tab(3).Control(12)=   "Line1"
            Tab(3).Control(13)=   "Label29"
            Tab(3).Control(14)=   "Label28"
            Tab(3).Control(15)=   "Label27"
            Tab(3).Control(16)=   "Label25"
            Tab(3).Control(17)=   "Label23"
            Tab(3).Control(18)=   "Label22"
            Tab(3).ControlCount=   19
            Begin VB.CommandButton Command1 
               Caption         =   "View Report"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   -74640
               TabIndex        =   51
               Top             =   5280
               Width           =   1215
            End
            Begin VB.Frame Frame5 
               Caption         =   "Period "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   -74880
               TabIndex        =   40
               Top             =   480
               Width           =   6375
               Begin VB.Label Label21 
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   3000
                  TabIndex        =   43
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label Label26 
                  Caption         =   "Label26"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   42
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label Label24 
                  Caption         =   "Label24"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   840
                  TabIndex        =   41
                  Top             =   360
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Work Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2655
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   6375
               Begin VB.PictureBox ListView2 
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   2175
                  Left            =   120
                  ScaleHeight     =   2115
                  ScaleWidth      =   6075
                  TabIndex        =   25
                  Top             =   360
                  Width           =   6135
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Period "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   240
               TabIndex        =   19
               Top             =   3240
               Width           =   6255
               Begin VB.Label Label1 
                  Caption         =   "Date Started: "
                  Height          =   375
                  Left            =   240
                  TabIndex        =   23
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Caption         =   "Date Ended:"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   22
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Left            =   1560
                  TabIndex        =   21
                  Top             =   360
                  Width           =   4455
               End
               Begin VB.Label Label10 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Left            =   1560
                  TabIndex        =   20
                  Top             =   840
                  Width           =   4455
               End
            End
            Begin VB.ComboBox cboDeductionType 
               Height          =   315
               Left            =   -73080
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   720
               Width           =   3135
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   -73080
               TabIndex        =   16
               Text            =   "0"
               Top             =   1200
               Width           =   3135
            End
            Begin VB.CommandButton Command2 
               Height          =   495
               Left            =   -69120
               Picture         =   "frmTransact.frx":4012C
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   2040
               Width           =   615
            End
            Begin VB.CommandButton Command3 
               Height          =   495
               Left            =   -69720
               Picture         =   "frmTransact.frx":40B2E
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   2040
               Width           =   615
            End
            Begin VB.Frame Frame4 
               Caption         =   "Allowances "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   -74760
               TabIndex        =   7
               Top             =   1440
               Width           =   4935
               Begin VB.ComboBox cboAllowance 
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.TextBox Text1 
                  Height          =   375
                  Left            =   1680
                  TabIndex        =   8
                  Text            =   "0"
                  Top             =   720
                  Width           =   2295
               End
               Begin VB.Label Label41 
                  Caption         =   "*"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   4080
                  TabIndex        =   62
                  Top             =   360
                  Width           =   135
               End
               Begin VB.Label Label16 
                  Caption         =   "Type of Allowance :"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   11
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.Label Label15 
                  Caption         =   "Amount : "
                  Height          =   375
                  Left            =   120
                  TabIndex        =   10
                  Top             =   720
                  Width           =   855
               End
            End
            Begin VB.CommandButton Command5 
               Height          =   495
               Left            =   -69720
               Picture         =   "frmTransact.frx":41530
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   2280
               Width           =   615
            End
            Begin VB.CommandButton Command4 
               Height          =   495
               Left            =   -69120
               Picture         =   "frmTransact.frx":41F32
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   2280
               Width           =   615
            End
            Begin VB.PictureBox ListView4 
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2295
               Left            =   -74760
               ScaleHeight     =   2235
               ScaleWidth      =   6075
               TabIndex        =   4
               Top             =   3480
               Width           =   6135
            End
            Begin VB.PictureBox DTPicker2 
               Height          =   375
               Left            =   -73080
               ScaleHeight     =   315
               ScaleWidth      =   3075
               TabIndex        =   13
               Top             =   2160
               Width           =   3135
            End
            Begin VB.PictureBox DTPicker1 
               Height          =   375
               Left            =   -73080
               ScaleHeight     =   315
               ScaleWidth      =   3075
               TabIndex        =   14
               Top             =   1680
               Width           =   3135
            End
            Begin VB.PictureBox ListView3 
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2535
               Left            =   -74760
               ScaleHeight     =   2475
               ScaleWidth      =   6075
               TabIndex        =   18
               Top             =   3240
               Width           =   6135
            End
            Begin VB.Line Line5 
               X1              =   -74760
               X2              =   -68520
               Y1              =   3000
               Y2              =   3000
            End
            Begin VB.Label Label43 
               Caption         =   "Required Field"
               Height          =   255
               Left            =   -74520
               TabIndex        =   64
               Top             =   3120
               Width           =   1455
            End
            Begin VB.Label Label42 
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   -74760
               TabIndex        =   63
               Top             =   3120
               Width           =   135
            End
            Begin VB.Line Line4 
               X1              =   -74760
               X2              =   -68520
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Label Label40 
               Caption         =   "Required Field"
               Height          =   255
               Left            =   -74520
               TabIndex        =   61
               Top             =   2880
               Width           =   1455
            End
            Begin VB.Label Label39 
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   -74760
               TabIndex        =   60
               Top             =   2880
               Width           =   135
            End
            Begin VB.Label Label38 
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   -69840
               TabIndex        =   59
               Top             =   720
               Width           =   135
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Label37"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   -70440
               TabIndex        =   58
               Top             =   4920
               Width           =   1695
            End
            Begin VB.Label Label36 
               Caption         =   "Label36"
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
               Left            =   -72840
               TabIndex        =   57
               Top             =   4200
               Width           =   1695
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Label35"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70440
               TabIndex        =   56
               Top             =   3720
               Width           =   1695
            End
            Begin VB.Label Label34 
               Caption         =   "Label34"
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
               Left            =   -72840
               TabIndex        =   55
               Top             =   3000
               Width           =   1695
            End
            Begin VB.Label Label33 
               Caption         =   "Label33"
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
               Left            =   -72840
               TabIndex        =   54
               Top             =   2640
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Label32"
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
               Left            =   -72840
               TabIndex        =   53
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label31 
               Caption         =   "Label31"
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
               Left            =   -72840
               TabIndex        =   52
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label Label30 
               Caption         =   "Net Salary : "
               Height          =   255
               Left            =   -71880
               TabIndex        =   50
               Top             =   4920
               Width           =   975
            End
            Begin VB.Line Line3 
               X1              =   -74760
               X2              =   -68640
               Y1              =   4680
               Y2              =   4680
            End
            Begin VB.Line Line2 
               X1              =   -74760
               X2              =   -68640
               Y1              =   3480
               Y2              =   3480
            End
            Begin VB.Line Line1 
               X1              =   -74760
               X2              =   -68640
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Label Label29 
               Caption         =   "Basic Salary : "
               Height          =   255
               Left            =   -74640
               TabIndex        =   49
               Top             =   2640
               Width           =   1095
            End
            Begin VB.Label Label28 
               Caption         =   "Gross Salary :"
               Height          =   375
               Left            =   -71880
               TabIndex        =   48
               Top             =   3720
               Width           =   1095
            End
            Begin VB.Label Label27 
               Caption         =   "Total Deduction : "
               Height          =   375
               Left            =   -74520
               TabIndex        =   47
               Top             =   4200
               Width           =   1575
            End
            Begin VB.Label Label25 
               Caption         =   "Total Allowance :"
               Height          =   375
               Left            =   -74640
               TabIndex        =   46
               Top             =   3000
               Width           =   1455
            End
            Begin VB.Label Label23 
               Caption         =   "Total Absent/Tardy : "
               Height          =   255
               Left            =   -74640
               TabIndex        =   45
               Top             =   2040
               Width           =   1935
            End
            Begin VB.Label Label22 
               Caption         =   "Total Worked Hour :"
               Height          =   375
               Left            =   -74640
               TabIndex        =   44
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Worked Hour(s) : "
               Height          =   375
               Left            =   360
               TabIndex        =   39
               Top             =   4920
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Absent/Tardy : "
               Height          =   375
               Left            =   360
               TabIndex        =   38
               Top             =   5400
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Type of Deduction :"
               Height          =   375
               Left            =   -74640
               TabIndex        =   37
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label6 
               Caption         =   "Amount : "
               Height          =   375
               Left            =   -74640
               TabIndex        =   36
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Effectivity Date : "
               Height          =   375
               Left            =   -74640
               TabIndex        =   35
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "Termination Date : "
               Height          =   375
               Left            =   -74640
               TabIndex        =   34
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1800
               TabIndex        =   33
               Top             =   4800
               Width           =   4455
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1800
               TabIndex        =   32
               Top             =   5400
               Width           =   4455
            End
            Begin VB.Label Label13 
               Caption         =   "Position    : "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   31
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label14 
               Height          =   375
               Left            =   -73680
               TabIndex        =   30
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label17 
               Caption         =   "Basic Salary : "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -71640
               TabIndex        =   29
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label18 
               Height          =   375
               Left            =   -70440
               TabIndex        =   28
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label19 
               Caption         =   "Amount of Exemption : "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   27
               Top             =   1080
               Width           =   1935
            End
            Begin VB.Label Label20 
               Height          =   375
               Left            =   -72600
               TabIndex        =   26
               Top             =   1080
               Width           =   1695
            End
         End
      End
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frmTransact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private str1, str2 As String
Private Sub Command2_Click()
On Error GoTo errtrap
Dim Increment As Integer
Call fConn
SQL "select * from tbldeduction;"
If Not RS.EOF Then
SQL "select (max(empdeductionid)+1) as incremented from tbldeduction;"
Increment = RS!incremented
Else
Increment = 1
End If
With frmTransact
SQL "Insert into tbldeduction values(" & Val(Trim((Left(.cboDeductionType, 3)))) & "," & Val(.Text5) & ",'" & .DTPicker1 & "','" & .DTPicker2 & "'," & Val(Trim((Left(.Frame1, 3)))) & "," & Increment & "," & Val(frmTransact.ListView2.SelectedItem.Text) & ")"
End With
MsgBox "Data saved."
frmTransact.ListView3.ListItems.Clear
SQL "Select * from tblDeduction where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by empdeductionid asc"
Do While Not RS.EOF
          With frmTransact.ListView3.ListItems
            Set Item = .Add(, , RS!empdeductionid)
              Item.SubItems(1) = RS!typeofdeduction
              Item.SubItems(2) = RS!deduction_amount
              Item.SubItems(3) = RS!effective_date
              Item.SubItems(4) = RS!termination_date
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
errtrap:
 Select Case Err.Number
    Case -2147217887
      MsgBox "Please fill up all the required fields.", vbCritical, "Error"
 '   Case Else
 '     MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Sub
Private Sub Command3_Click()
On Error GoTo errtrap
If frmTransact.ListView3.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tbldeduction WHERE empdeductionid = " & Val(frmTransact.ListView3.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
frmTransact.ListView3.ListItems.Clear
SQL "Select * from tbldeduction where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by empdeductionid asc"
Do While Not RS.EOF
          With frmTransact.ListView3.ListItems
            Set Item = .Add(, , RS!empdeductionid)
              Item.SubItems(1) = RS!typeofdeduction
              Item.SubItems(2) = RS!deduction_amount
              Item.SubItems(3) = RS!effective_date
              Item.SubItems(4) = RS!termination_date
          End With
          RS.MoveNext
        Loop
    Conn.Close
    Set Conn = Nothing
errtrap:
End Sub
Private Sub Command4_Click()
On Error GoTo errtrap
Dim Increment As Integer
Call fConn
SQL "select * from tblallowance;"
If Not RS.EOF Then
SQL "select (max(allowanceid)+1) as incremented from tblallowance;"
Increment = RS!incremented
Else
Increment = 1
End If
With frmTransact
SQL "Insert into tblallowance values(" & Increment & "," & Val(Trim((Left(.cboAllowance, 3)))) & "," & Val(.Text1) & "," & Val(Trim((Left(.Frame1, 3)))) & "," & Val(frmTransact.ListView2.SelectedItem.Text) & ")"
End With
MsgBox "Data saved."
frmTransact.ListView4.ListItems.Clear
'MsgBox (Val(Trim((Left(Me.Frame1, 3)))))
'SQL "Select * from tblallowance where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " order by allowanceid asc"
'MsgBox (RS!allowanceid)
'MsgBox (RS!type_allowance)
'MsgBox (RS!allow_amount)
       'Do While Not RS.EOF
        'With frmTransact.ListView4.ListItems
            'Set Item = .Add(, , RS!allowanceid)
             'Item.SubItems(1) = RS!type_allowance
           '   Item.SubItems(2) = RS!allow_amount
          'End With
          'RS.MoveNext
        'Loop
SQL "Select * from tblallowance where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by allowanceid asc "
Do While Not RS.EOF
          With frmTransact.ListView4.ListItems
            Set Item = .Add(, , RS!allowanceid)
              Item.SubItems(1) = RS!type_allowanceid
              Item.SubItems(2) = RS!allow_amount
          End With
          RS.MoveNext
        Loop

Conn.Close
Set Conn = Nothing
errtrap:
 Select Case Err.Number
    Case -2147217887
      MsgBox "Please fill up all the required fields.", vbCritical, "Error"
 '   Case Else
 '     MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Sub

Private Sub Command5_Click()
On Error GoTo errtrap

If frmTransact.ListView4.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tblallowance WHERE allowanceid = " & Val(frmTransact.ListView4.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
Call fConn
frmTransact.ListView4.ListItems.Clear
SQL "Select * from tblallowance where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by allowanceid asc"
Do While Not RS.EOF
          With frmTransact.ListView4.ListItems
            Set Item = .Add(, , RS!allowanceid)
              Item.SubItems(1) = RS!type_allowanceid
              Item.SubItems(2) = RS!allow_amount
          End With
          RS.MoveNext
        Loop

Conn.Close
Set Conn = Nothing
errtrap:
End Sub

Private Sub Form_Load()
Me.SSTab1.Tab = 0

Call fConn
    SQL "Select * from tblemployee order by employeeid asc"
        Do While Not RS.EOF
          With frmTransact.ListView1.ListItems
            Set Item = .Add(, , RS!employeeid)
              Item.SubItems(1) = StrConv(RS!lastname, vbProperCase)
              Item.SubItems(2) = StrConv(RS!firstname, vbProperCase)
              Item.SubItems(3) = StrConv(Right(RS!middlename, 1), vbProperCase)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
End Sub



Private Sub ListView1_Click()
On Error GoTo errtrap

Me.Frame1.Enabled = True
frmTransact.ListView2.ListItems.Clear
Call fConn
    SQL "Select * from tblemployee where employeeid = " & Val(frmTransact.ListView1.SelectedItem.Text) & ""
frmTransact.Frame1.Caption = RS!employeeid & Space(3) & UCase(RS!lastname) & Space(1) & "," & Space(1) & UCase(RS!firstname) & Space(2) & UCase(Right(RS!middlename, 1))
    SQL "Select * from tblemp_attendance where employeeid = " & Val(frmTransact.ListView1.SelectedItem.Text) & "  order by attendanceid asc"
        If Not RS.EOF Then
        Me.Framex.Enabled = True
        Do While Not RS.EOF
          With frmTransact.ListView2.ListItems
            Set Item = .Add(, , RS!attendanceid)
              Item.SubItems(1) = RS!datestarted
              Item.SubItems(2) = RS!dateended
              Item.SubItems(3) = RS!workedhours
              Item.SubItems(4) = RS!absent_tardy
          End With
          RS.MoveNext
        Loop
        
    SQL "Select * from tblemp_attendance where attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & ""
        Do While Not RS.EOF
              Label9 = RS!datestarted
              Label10 = RS!dateended
              Label11 = RS!workedhours
              Label12 = RS!absent_tardy
              
              Label24 = RS!datestarted
              Label26 = RS!dateended
          RS.MoveNext
        Loop
        
  Conn.Close
  Set Conn = Nothing
  Else
  Me.SSTab1.Tab = 0
  Me.Framex.Enabled = False
  End If
  Call rushCal
errtrap:
  Select Case Err.Number
    Case 91
      MsgBox "No record found.", vbCritical, "Error"
    'Case Else
      'MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Sub

Private Sub ListView2_Click()
Call fConn
    SQL "Select * from tblemp_attendance where attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & ""
        Do While Not RS.EOF
              Label9 = RS!datestarted
              Label10 = RS!dateended
              Label11 = RS!workedhours
              Label12 = RS!absent_tardy
              
              Label24 = RS!datestarted
              Label26 = RS!dateended
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
  Call rushCal
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
Dim tmpID As Integer

On Error GoTo errtrap

If Me.SSTab1.Tab = 0 Then
Me.cboDeductionType.Clear
Call deductionTypeList
ElseIf Me.SSTab1.Tab = 1 Then
Call fConn
frmTransact.ListView3.ListItems.Clear

SQL "Select * from tbldeduction where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by empdeductionid asc"
Do While Not RS.EOF
          With frmTransact.ListView3.ListItems
            Set Item = .Add(, , RS!empdeductionid)
              Item.SubItems(1) = RS!typeofdeduction
              Item.SubItems(2) = RS!deduction_amount
              Item.SubItems(3) = RS!effective_date
              Item.SubItems(4) = RS!termination_date
          End With
          RS.MoveNext
        Loop
    Conn.Close
    Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 2 Then
Me.cboAllowance.Clear
Call allowanceTypeList
Call fConn
SQL "Select * from tblemployee where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & ""
tmpID = Val(RS!emposition)
SQL "Select * from tblposition where positionid = " & tmpID & ""
Me.Label14 = StrConv(RS!position_name, vbProperCase)
Me.Label18 = RS!basic_salary
Me.Label20 = RS!amount_exemption

frmTransact.ListView4.ListItems.Clear
'Conn.Close
'Set Conn = Nothing
'Call fConn
'MsgBox (Val(frmTransact.ListView2.SelectedItem.Text))
'MsgBox (Val(Trim((Left(Me.Frame1, 3)))))
SQL "Select * from tblallowance where employeeid = " & Val(Trim((Left(Me.Frame1, 3)))) & " AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " order by allowanceid asc"
'MsgBox (RS!attendaceid)
'AND attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & "
Do While Not RS.EOF
          With frmTransact.ListView4.ListItems
            Set Item = .Add(, , RS!allowanceid)
              Item.SubItems(1) = RS!type_allowanceid
              Item.SubItems(2) = RS!allow_amount
          End With
          RS.MoveNext
        Loop
Conn.Close
Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 4 Then
Me.Label31.Caption = Me.Label11.Caption
End If
errtrap:
End Sub

Function rushCal()
Dim total_deduction, total_allowance As Double
Call fConn
'SQL "Select * from tblemp_attendance,tblposition,tblallowance,tbldeduction,tblemployee where tblemployee.employeeid = " & 2 & ""
 SQL "Select workedhours,absent_tardy,basic_salary from tblemp_attendance join tblemployee on(tblemployee.employeeid = tblemp_attendance.employeeid) join tblposition on (tblemployee.emposition = tblposition.positionid) where tblemployee.employeeid = " & Val(frmTransact.ListView1.SelectedItem.Text) & " AND tblemp_attendance.attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " "
 'Do While Not RS.EOF
 If Not RS.EOF Then
 Me.Label31.Caption = RS!workedhours
 Me.Label32.Caption = RS!absent_tardy
 Me.Label33.Caption = RS!basic_salary
 'RS.MoveNext
 'Loop
 End If
 SQL " Select deduction_amount from tbldeduction where employeeid = " & Val(frmTransact.ListView1.SelectedItem.Text) & " and attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " "
 Do While Not RS.EOF
 total_deduction = total_deduction + Val(RS!deduction_amount)
 RS.MoveNext
 Loop
 Me.Label36 = total_deduction
 
 SQL " Select allow_amount from tblallowance where employeeid = " & Val(frmTransact.ListView1.SelectedItem.Text) & " and attendanceid = " & Val(frmTransact.ListView2.SelectedItem.Text) & " "
 Do While Not RS.EOF
 total_allowance = total_allowance + Val(RS!allow_amount)
 RS.MoveNext
 Loop
 Me.Label34 = total_allowance
 Me.Label35 = Val(Label33) + total_allowance
 Me.Label37 = FormatNumber((Val(Label35) - total_deduction - (Val(Label32) * (Val(Label33) / 22))), 2, True, True, True)
 Conn.Close
Set Conn = Nothing
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub
