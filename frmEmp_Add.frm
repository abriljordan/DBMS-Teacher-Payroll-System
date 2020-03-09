VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmp_Add 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emloyee Registration"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEmp_Add.frx":0000
   ScaleHeight     =   6870
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save and Close"
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
      Left            =   1320
      TabIndex        =   52
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ca&ncel"
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
      Left            =   4440
      TabIndex        =   30
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
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
      Left            =   3240
      TabIndex        =   29
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmp_Add 
      Caption         =   "Save"
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
      Left            =   120
      TabIndex        =   28
      Top             =   6240
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmEmp_Add.frx":400BC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(8)=   "Line2"
      Tab(0).Control(9)=   "Label23"
      Tab(0).Control(10)=   "txtLastName"
      Tab(0).Control(11)=   "txtFirstName"
      Tab(0).Control(12)=   "txtMiddleName"
      Tab(0).Control(13)=   "txtTIN"
      Tab(0).Control(14)=   "DTPicker1"
      Tab(0).Control(15)=   "DTPicker2"
      Tab(0).Control(16)=   "cboGender"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Additional Information"
      TabPicture(1)   =   "frmEmp_Add.frx":400D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "txtNotes"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Registration"
      TabPicture(2)   =   "frmEmp_Add.frx":400F4
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label25"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label26"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboCivilStatus"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtDependent"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.ComboBox cboGender 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Frame Frame4 
         Caption         =   "Location "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   44
         Top             =   3600
         Width           =   5175
         Begin VB.ComboBox cboRegion 
            Height          =   315
            Left            =   1080
            TabIndex        =   50
            Top             =   1440
            Width           =   3735
         End
         Begin VB.ComboBox cboDivision 
            Height          =   315
            Left            =   1080
            TabIndex        =   48
            Top             =   960
            Width           =   3735
         End
         Begin VB.ComboBox cboStation 
            Height          =   315
            Left            =   1080
            TabIndex        =   46
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label29 
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
            Left            =   4920
            TabIndex        =   58
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label28 
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
            Left            =   4920
            TabIndex        =   57
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label27 
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
            Left            =   4920
            TabIndex        =   56
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label22 
            Caption         =   "Region:"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Division:"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "Station:"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox txtDependent 
         Height          =   405
         Left            =   1200
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cboCivilStatus 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Registration "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   5175
         Begin VB.ComboBox cboStep 
            Height          =   315
            Left            =   3120
            TabIndex        =   43
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cboGrade 
            Height          =   315
            Left            =   1080
            TabIndex        =   41
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cboPosition 
            Height          =   315
            Left            =   1080
            TabIndex        =   35
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label24 
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
            Left            =   4920
            TabIndex        =   53
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label19 
            Caption         =   "Step:"
            Height          =   375
            Left            =   2640
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "Grade:"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Position:"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.PictureBox DTPicker2 
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
         Left            =   -73680
         ScaleHeight     =   315
         ScaleWidth      =   2955
         TabIndex        =   32
         Top             =   4440
         Width           =   3015
      End
      Begin VB.PictureBox DTPicker1 
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
         Left            =   -73680
         ScaleHeight     =   315
         ScaleWidth      =   2955
         TabIndex        =   31
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txtTIN 
         Height          =   375
         Left            =   -73680
         TabIndex        =   24
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox txtNotes 
         Height          =   1365
         Left            =   -74760
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   4560
         Width           =   4935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Employee Address "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   5175
         Begin VB.TextBox txtStreet 
            Height          =   375
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtProvince 
            Height          =   375
            Left            =   1320
            TabIndex        =   17
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox txtRegion 
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label6 
            Caption         =   "Street/Brngy:"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Province/City:"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Region:"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contact "
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
         Left            =   -74880
         TabIndex        =   10
         Top             =   2520
         Width           =   5175
         Begin VB.TextBox txtPhone 
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "Cell/Phone:"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "E-mail:"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   375
         Left            =   -73680
         TabIndex        =   9
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtFirstName 
         Height          =   375
         Left            =   -73680
         TabIndex        =   7
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtLastName 
         Height          =   375
         Left            =   -73680
         TabIndex        =   6
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label26 
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
         Left            =   120
         TabIndex        =   55
         Top             =   5760
         Width           =   135
      End
      Begin VB.Label Label25 
         Caption         =   "Required Field"
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Gender:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   51
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Dependent:"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Civil Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   -74880
         X2              =   -69720
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label14 
         Caption         =   "Employment Date:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   27
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Date of Birth:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "SIN/SNN:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   -72840
         X2              =   -69720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label11 
         Caption         =   "Notes (max 255 characters)"
         Height          =   375
         Left            =   -74880
         TabIndex        =   22
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Middle Name:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "First Name:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -73680
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Employee ID:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEmp_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboGrade_Change()
If ((Val(cboGrade.Text) < 11) And (Val(cboGrade.Text) > 0)) Then
'ok
Else
cboGrade.Text = ""
End If
End Sub
Private Sub cboStep_Change()
If ((Val(cboStep.Text) < 11) And (Val(cboStep.Text) > 0)) Then
'ok
Else
cboStep.Text = ""
End If
End Sub

Private Sub cmdEmp_Add_Click()
Call addEmployee
Call loadRecords
Call loadEmpID
Call clearField
End Sub

Private Sub Command1_Click()
Call addEmployee
Call loadRecords
Unload Me
End Sub

Private Sub Command2_Click()
Call clearField
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo errtrap
If frmMenu.ListView1.SelectedItem = frmMenu.ListView1.ListItems(1) Then
Call loadEmpID
End If
errtrap:
End Sub

Private Sub Form_Load()
Me.SSTab1.Tab = 0
  With Me.cboCivilStatus
    .AddItem "Single"
    .AddItem "Married"
    .AddItem "Widow"
    .AddItem "Separate"
  End With
Me.cboGender.AddItem "Male"
Me.cboGender.AddItem "Female"
  With Me.cboGrade
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
  End With
    With Me.cboStep
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
  End With
Call stationList
Call divisionList
Call regionList
Call positionList
End Sub


Private Sub txtDependent_KeyPress(KeyAscii As Integer)
    KeyAscii = DigitOnly(KeyAscii)
End Sub
