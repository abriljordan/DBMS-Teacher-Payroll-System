VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmpSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEmpSetting.frx":0000
   ScaleHeight     =   4395
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Position     "
      TabPicture(0)   =   "frmEmpSetting.frx":400BC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ListView1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Station     "
      TabPicture(1)   =   "frmEmpSetting.frx":400D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(4)=   "ListView2"
      Tab(1).Control(5)=   "Command7"
      Tab(1).Control(6)=   "Command8"
      Tab(1).Control(7)=   "Command9"
      Tab(1).Control(8)=   "Command10"
      Tab(1).Control(9)=   "Command11"
      Tab(1).Control(10)=   "Command12"
      Tab(1).Control(11)=   "Text4"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Division     "
      TabPicture(2)   =   "frmEmpSetting.frx":400F4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(3)=   "Line3"
      Tab(2).Control(4)=   "ListView3"
      Tab(2).Control(5)=   "Text3"
      Tab(2).Control(6)=   "Command13"
      Tab(2).Control(7)=   "Command14"
      Tab(2).Control(8)=   "Command15"
      Tab(2).Control(9)=   "Command16"
      Tab(2).Control(10)=   "Command17"
      Tab(2).Control(11)=   "Command18"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Region     "
      TabPicture(3)   =   "frmEmpSetting.frx":40110
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line4"
      Tab(3).Control(1)=   "Label11"
      Tab(3).Control(2)=   "Label12"
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(4)=   "ListView4"
      Tab(3).Control(5)=   "Command19"
      Tab(3).Control(6)=   "Command20"
      Tab(3).Control(7)=   "Command21"
      Tab(3).Control(8)=   "Command22"
      Tab(3).Control(9)=   "Command23"
      Tab(3).Control(10)=   "Command24"
      Tab(3).Control(11)=   "Text5"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Allowance  "
      TabPicture(4)   =   "frmEmpSetting.frx":4012C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label15"
      Tab(4).Control(1)=   "Label16"
      Tab(4).Control(2)=   "Label17"
      Tab(4).Control(3)=   "Line5"
      Tab(4).Control(4)=   "ListView5"
      Tab(4).Control(5)=   "Text7"
      Tab(4).Control(6)=   "Command25"
      Tab(4).Control(7)=   "Command26"
      Tab(4).Control(8)=   "Command27"
      Tab(4).Control(9)=   "Command28"
      Tab(4).Control(10)=   "Command29"
      Tab(4).Control(11)=   "Command30"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "Deduction  "
      TabPicture(5)   =   "frmEmpSetting.frx":40148
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label18"
      Tab(5).Control(1)=   "Label19"
      Tab(5).Control(2)=   "Label20"
      Tab(5).Control(3)=   "Line6"
      Tab(5).Control(4)=   "ListView6"
      Tab(5).Control(5)=   "Text8"
      Tab(5).Control(6)=   "Command31"
      Tab(5).Control(7)=   "Command32"
      Tab(5).Control(8)=   "Command33"
      Tab(5).Control(9)=   "Command34"
      Tab(5).Control(10)=   "Command35"
      Tab(5).Control(11)=   "Command36"
      Tab(5).ControlCount=   12
      Begin VB.CommandButton Command36 
         Caption         =   ">>|"
         Height          =   375
         Left            =   -67560
         TabIndex        =   66
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command35 
         Caption         =   ">"
         Height          =   375
         Left            =   -68280
         TabIndex        =   65
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command34 
         Caption         =   "<"
         Height          =   375
         Left            =   -69000
         TabIndex        =   64
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command33 
         Caption         =   "|<<"
         Height          =   375
         Left            =   -69720
         TabIndex        =   63
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71640
         TabIndex        =   62
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72840
         TabIndex        =   61
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   -73440
         TabIndex        =   60
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command30 
         Caption         =   ">>|"
         Height          =   375
         Left            =   -67560
         TabIndex        =   55
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command29 
         Caption         =   ">"
         Height          =   375
         Left            =   -68280
         TabIndex        =   54
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command28 
         Caption         =   "<"
         Height          =   375
         Left            =   -69000
         TabIndex        =   53
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command27 
         Caption         =   "|<<"
         Height          =   375
         Left            =   -69720
         TabIndex        =   52
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71640
         TabIndex        =   51
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72840
         TabIndex        =   50
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   -73440
         TabIndex        =   49
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1560
         TabIndex        =   47
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   -73440
         TabIndex        =   42
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72840
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71640
         TabIndex        =   40
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command22 
         Caption         =   "|<<"
         Height          =   375
         Left            =   -69720
         TabIndex        =   39
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command21 
         Caption         =   "<"
         Height          =   375
         Left            =   -69000
         TabIndex        =   38
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command20 
         Caption         =   ">"
         Height          =   375
         Left            =   -68280
         TabIndex        =   37
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command19 
         Caption         =   ">>|"
         Height          =   375
         Left            =   -67560
         TabIndex        =   36
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Caption         =   ">>|"
         Height          =   375
         Left            =   -67560
         TabIndex        =   31
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command17 
         Caption         =   ">"
         Height          =   375
         Left            =   -68280
         TabIndex        =   30
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command16 
         Caption         =   "<"
         Height          =   375
         Left            =   -69000
         TabIndex        =   29
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command15 
         Caption         =   "|<<"
         Height          =   375
         Left            =   -69720
         TabIndex        =   28
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71640
         TabIndex        =   27
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72840
         TabIndex        =   26
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   -73440
         TabIndex        =   25
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   -73440
         TabIndex        =   21
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -72840
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71640
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "|<<"
         Height          =   375
         Left            =   -69720
         TabIndex        =   18
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "<"
         Height          =   375
         Left            =   -69000
         TabIndex        =   17
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">"
         Height          =   375
         Left            =   -68280
         TabIndex        =   16
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">>|"
         Height          =   375
         Left            =   -67560
         TabIndex        =   15
         Top             =   3360
         Width           =   735
      End
      Begin VB.PictureBox ListView1 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   4320
         ScaleHeight     =   2595
         ScaleWidth      =   4755
         TabIndex        =   13
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">>|"
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">"
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "|<<"
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.PictureBox ListView2 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   -69720
         ScaleHeight     =   2595
         ScaleWidth      =   3795
         TabIndex        =   14
         Top             =   600
         Width           =   3855
      End
      Begin VB.PictureBox ListView3 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   -69720
         ScaleHeight     =   2595
         ScaleWidth      =   3555
         TabIndex        =   32
         Top             =   600
         Width           =   3615
      End
      Begin VB.PictureBox ListView4 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   -69720
         ScaleHeight     =   2595
         ScaleWidth      =   3675
         TabIndex        =   43
         Top             =   600
         Width           =   3735
      End
      Begin VB.PictureBox ListView5 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   -69720
         ScaleHeight     =   2595
         ScaleWidth      =   3675
         TabIndex        =   56
         Top             =   600
         Width           =   3735
      End
      Begin VB.PictureBox ListView6 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   -69720
         ScaleHeight     =   2595
         ScaleWidth      =   3675
         TabIndex        =   67
         Top             =   600
         Width           =   3735
      End
      Begin VB.Line Line6 
         X1              =   -74760
         X2              =   -70080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label20 
         Caption         =   "Region Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   70
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Label2"
         Height          =   375
         Left            =   -73440
         TabIndex        =   69
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Region ID:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   68
         Top             =   840
         Width           =   975
      End
      Begin VB.Line Line5 
         X1              =   -74760
         X2              =   -70080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label17 
         Caption         =   "Region Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   59
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Label2"
         Height          =   375
         Left            =   -73440
         TabIndex        =   58
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Region ID:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   57
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Amount Exemption:"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Region ID:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Label2"
         Height          =   375
         Left            =   -73440
         TabIndex        =   45
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Region Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Line Line4 
         X1              =   -74760
         X2              =   -70080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   -74760
         X2              =   -70080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label10 
         Caption         =   "Division Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
         Height          =   375
         Left            =   -73440
         TabIndex        =   34
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Division ID:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Station ID:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label2"
         Height          =   375
         Left            =   -73440
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Station Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   -74760
         X2              =   -70080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   4200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "Basic Salary:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Position Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Position ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEmpSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView1.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tblposition WHERE positionid = " & Val(frmEmpSetting.ListView1.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView1.ListItems.Clear
SQL "Select * from tblposition order by positionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView1.ListItems
            Set Item = .Add(, , RS!positionid)
              Item.SubItems(1) = RS!position_name
              Item.SubItems(2) = RS!basic_salary
              Item.SubItems(3) = RS!amount_exemption
          End With
          RS.MoveNext
        Loop
SQL "select (max(positionid)+1) as incremented from tblposition;"
Increment = RS!incremented
Me.Label2.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub
Private Sub Command11_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tblstation values(" & Val(.Label7) & ",'" & LCase(.Text4) & "')"
End With
MsgBox "Data saved."
frmEmpSetting.ListView2.ListItems.Clear
SQL "Select * from tblstation order by stationid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView2.ListItems
            Set Item = .Add(, , RS!stationid)
              Item.SubItems(1) = RS!station_name
          End With
          RS.MoveNext
        Loop
Me.Text4.Text = ""
SQL "select (max(stationid)+1) as incremented from tblstation;"
Increment = RS!incremented
Me.Label7.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub
Private Sub Command12_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView2.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tblstation WHERE stationid = " & Val(frmEmpSetting.ListView2.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView2.ListItems.Clear
SQL "Select * from tblstation order by stationid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView2.ListItems
            Set Item = .Add(, , RS!stationid)
              Item.SubItems(1) = RS!station_name
          End With
          RS.MoveNext
        Loop
SQL "select (max(stationid)+1) as incremented from tblstation;"
Increment = RS!incremented
Me.Label7.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub
Private Sub Command13_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView3.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tbldivision WHERE divisionid = " & Val(frmEmpSetting.ListView3.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView3.ListItems.Clear
SQL "Select * from tbldivision order by divisionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView3.ListItems
            Set Item = .Add(, , RS!divisionid)
              Item.SubItems(1) = RS!division_name
          End With
          RS.MoveNext
        Loop
SQL "select (max(divisionid)+1) as incremented from tbldivision;"
Increment = RS!incremented
Me.Label9.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub
Private Sub Command14_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tbldivision values(" & Val(.Label9) & ",'" & LCase(.Text3) & "')"
End With
MsgBox "Data saved."
frmEmpSetting.ListView3.ListItems.Clear
SQL "Select * from tbldivision order by divisionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView3.ListItems
            Set Item = .Add(, , RS!divisionid)
              Item.SubItems(1) = RS!division_name
          End With
          RS.MoveNext
        Loop
Me.Text3.Text = ""
SQL "select (max(divisionid)+1) as incremented from tbldivision;"
Increment = RS!incremented
Me.Label9.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub
Private Sub Command23_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tblregion values(" & Val(.Label12) & ",'" & LCase(.Text5) & "')"
End With
MsgBox "Data saved."
frmEmpSetting.ListView4.ListItems.Clear
SQL "Select * from tblregion order by regionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView4.ListItems
            Set Item = .Add(, , RS!regionid)
              Item.SubItems(1) = UCase(RS!region_name)
          End With
          RS.MoveNext
        Loop
Me.Text5.Text = ""
SQL "select (max(regionid)+1) as incremented from tblregion;"
Increment = RS!incremented
Me.Label12.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub

Private Sub Command24_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView4.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tblregion WHERE regionid = " & Val(frmEmpSetting.ListView4.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView4.ListItems.Clear
SQL "Select * from tblregion order by regionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView4.ListItems
            Set Item = .Add(, , RS!regionid)
              Item.SubItems(1) = UCase(RS!region_name)
          End With
          RS.MoveNext
        Loop
SQL "select (max(regionid)+1) as incremented from tblregion;"
Increment = RS!incremented
Me.Label12.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub

Private Sub Command25_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView5.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tbltypeofallowance WHERE type_allowanceid = " & Val(frmEmpSetting.ListView5.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView5.ListItems.Clear
SQL "Select * from tbltypeofallowance order by type_allowanceid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView5.ListItems
            Set Item = .Add(, , RS!type_allowanceid)
              Item.SubItems(1) = UCase(RS!allowance_descrip)
          End With
          RS.MoveNext
        Loop
SQL "select (max(type_allowanceid)+1) as incremented from tbltypeofallowance;"
Increment = RS!incremented
Me.Label16.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub
Private Sub Command26_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tbltypeofallowance values(" & Val(.Label16) & ",'" & LCase(.Text7) & "')"
End With
MsgBox "Data saved."
frmEmpSetting.ListView5.ListItems.Clear
SQL "Select * from tbltypeofallowance order by type_allowanceid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView5.ListItems
            Set Item = .Add(, , RS!type_allowanceid)
              Item.SubItems(1) = UCase(RS!allowance_descrip)
          End With
          RS.MoveNext
        Loop
Me.Text7.Text = ""
SQL "select (max(type_allowanceid)+1) as incremented from tbltypeofallowance;"
Increment = RS!incremented
Me.Label16.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub
Private Sub Command3_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tblposition values(" & Val(.Label2) & ",'" & LCase(.Text1) & "'," & Val(.Text2) & "," & Val(.Text6) & ")"
End With
MsgBox "Data saved."
frmEmpSetting.ListView1.ListItems.Clear
SQL "Select * from tblposition order by positionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView1.ListItems
            Set Item = .Add(, , RS!positionid)
              Item.SubItems(1) = RS!position_name
              Item.SubItems(2) = RS!basic_salary
              Item.SubItems(3) = RS!amount_exemption
          End With
          RS.MoveNext
        Loop
Me.Text1.Text = ""
Me.Text2.Text = ""
Me.Text6.Text = ""
SQL "select (max(positionid)+1) as incremented from tblposition;"
Increment = RS!incremented
Me.Label2.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub

Private Sub Command31_Click()
On Error GoTo errtrap
If frmEmpSetting.ListView6.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tbltypeofdeduction WHERE deductionid = " & Val(frmEmpSetting.ListView6.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
  Call fConn
  frmEmpSetting.ListView6.ListItems.Clear
SQL "Select * from tbltypeofdeduction order by deductionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView6.ListItems
            Set Item = .Add(, , RS!deductionid)
              Item.SubItems(1) = UCase(RS!deductionname)
          End With
          RS.MoveNext
        Loop
SQL "select (max(deductionid)+1) as incremented from tbltypeofdeduction;"
Increment = RS!incremented
Me.Label19.Caption = Increment
Conn.Close
Set Conn = Nothing
errtrap:
End Sub

Private Sub Command32_Click()
Call fConn
With frmEmpSetting
SQL "Insert into tbltypeofdeduction values(" & Val(.Label19) & ",'" & LCase(.Text8) & "')"
End With
MsgBox "Data saved."
frmEmpSetting.ListView6.ListItems.Clear
SQL "Select * from tbltypeofdeduction order by deductionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView6.ListItems
            Set Item = .Add(, , RS!deductionid)
              Item.SubItems(1) = UCase(RS!deductionname)
          End With
          RS.MoveNext
        Loop
Me.Text8.Text = ""
SQL "select (max(deductionid)+1) as incremented from tbltypeofdeduction;"
Increment = RS!incremented
Me.Label19.Caption = Increment
Conn.Close
Set Conn = Nothing
End Sub

Private Sub Form_Load()
Dim Increment As Integer
 Me.SSTab1.Tab = 0
 Command1.Enabled = False
 frmEmpSetting.ListView1.ListItems.Clear
Call fConn
SQL "Select * from tblposition order by positionid asc"
    'RS.MoveFirst
    Do While Not RS.EOF
          With frmEmpSetting.ListView1.ListItems
            Set Item = .Add(, , RS!positionid)
              Item.SubItems(1) = StrConv(RS!position_name, vbProperCase)
              Item.SubItems(2) = RS!basic_salary
              Item.SubItems(3) = RS!amount_exemption
          End With
          RS.MoveNext
        Loop
SQL "select (max(positionid)+1) as incremented from tblposition;"
Increment = RS!incremented
Me.Label2.Caption = Increment
  Conn.Close
  Set Conn = Nothing
End Sub
Private Sub ListView1_Click()
If frmEmpSetting.ListView1.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command1.Enabled = True
End If
End Sub
Private Sub ListView2_Click()
If frmEmpSetting.ListView2.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command12.Enabled = True
End If
End Sub
Private Sub ListView3_Click()
If frmEmpSetting.ListView3.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command13.Enabled = True
End If
End Sub
Private Sub ListView4_Click()
If frmEmpSetting.ListView4.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command24.Enabled = True
End If
End Sub
Private Sub ListView5_Click()
If frmEmpSetting.ListView5.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command25.Enabled = True
End If
End Sub
Private Sub ListView6_Click()
If frmEmpSetting.ListView6.ListItems.Count = 0 Then
    MsgBox "No records found.!", vbExclamation, "Error"
Else
 Command31.Enabled = True
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 0 Then
frmEmpSetting.ListView1.ListItems.Clear
Call fConn
SQL "Select * from tblposition order by positionid asc"
    'RS.MoveFirst
    Do While Not RS.EOF
          With frmEmpSetting.ListView1.ListItems
            Set Item = .Add(, , RS!positionid)
              Item.SubItems(1) = StrConv(RS!position_name, vbProperCase)
              Item.SubItems(2) = RS!basic_salary
              Item.SubItems(3) = RS!amount_exemption
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 1 Then
frmEmpSetting.ListView2.ListItems.Clear
Command12.Enabled = False
Call fConn
SQL "select (max(stationid)+1) as incremented from tblstation;"
Increment = RS!incremented
Me.Label7.Caption = Increment
Call fConn
SQL "Select * from tblstation order by stationid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView2.ListItems
            Set Item = .Add(, , RS!stationid)
              Item.SubItems(1) = StrConv(RS!station_name, vbProperCase)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 2 Then
frmEmpSetting.ListView3.ListItems.Clear
Command13.Enabled = False
Call fConn
SQL "select (max(divisionid)+1) as incremented from tbldivision;"
Increment = RS!incremented
Me.Label9.Caption = Increment
Call fConn
SQL "Select * from tbldivision order by divisionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView3.ListItems
            Set Item = .Add(, , RS!divisionid)
              Item.SubItems(1) = StrConv(RS!division_name, vbProperCase)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 3 Then
frmEmpSetting.ListView4.ListItems.Clear
Command24.Enabled = False
Call fConn
SQL "select (max(regionid)+1) as incremented from tblregion;"
Increment = RS!incremented
Me.Label12.Caption = Increment
Call fConn
SQL "Select * from tblregion order by regionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView4.ListItems
            Set Item = .Add(, , RS!regionid)
              Item.SubItems(1) = UCase(RS!region_name)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 4 Then
frmEmpSetting.ListView5.ListItems.Clear
Command25.Enabled = False
Call fConn
SQL "select (max(type_allowanceid)+1) as incremented from tbltypeofallowance;"
Increment = RS!incremented
Me.Label16.Caption = Increment
Call fConn
SQL "Select * from tbltypeofallowance order by type_allowanceid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView5.ListItems
            Set Item = .Add(, , RS!type_allowanceid)
              Item.SubItems(1) = UCase(RS!allowance_descrip)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
ElseIf Me.SSTab1.Tab = 5 Then
frmEmpSetting.ListView6.ListItems.Clear
Command31.Enabled = False
Call fConn
SQL "select (max(deductionid)+1) as incremented from tbltypeofdeduction;"
Increment = RS!incremented
Me.Label19.Caption = Increment
Call fConn
SQL "Select * from tbltypeofdeduction order by deductionid asc"
    Do While Not RS.EOF
          With frmEmpSetting.ListView6.ListItems
            Set Item = .Add(, , RS!deductionid)
              Item.SubItems(1) = UCase(RS!deductionname)
          End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub
