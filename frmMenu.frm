VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000003&
   Caption         =   "Teachers Payroll System"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ImageList2 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   1320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   26
      Top             =   4440
      Width           =   1200
   End
   Begin VB.PictureBox ListView5 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   9120
      ScaleHeight     =   5835
      ScaleWidth      =   5715
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.PictureBox ListView4 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3120
      ScaleHeight     =   5835
      ScaleWidth      =   5715
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   1320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   27
      Top             =   3000
      Width           =   1200
   End
   Begin VB.PictureBox ListView3 
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   2595
      TabIndex        =   11
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   1575
      Left            =   3120
      TabIndex        =   9
      Top             =   8400
      Width           =   11775
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
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
         Left            =   9240
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
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
         Left            =   10560
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000003&
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   11535
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Search : "
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   120
      ScaleHeight     =   4905
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   " Quick Launch  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton Command2 
         Caption         =   "Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6000
         Picture         =   "frmMenu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7920
         Picture         =   "frmMenu.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEmpSetting 
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         Picture         =   "frmMenu.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEmpAttendance 
         Caption         =   "Employee Attendance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2160
         Picture         =   "frmMenu.frx":29CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEmployee 
         Caption         =   "Employees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         Picture         =   "frmMenu.frx":3898
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   0
      Width           =   15135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   10200
      Width           =   15135
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Server : "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         TabIndex        =   22
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Localhost"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13800
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User : "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox ListView2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3120
      ScaleHeight     =   5835
      ScaleWidth      =   11715
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   18000
      Left            =   -4200
      Picture         =   "frmMenu.frx":4162
      Top             =   1560
      Width           =   28800
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   480
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   480
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1320
      Top             =   8160
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   1800
      Y2              =   10080
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
If MsgBox("Are you sure you want to exit.  Continue?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
End If
End Sub

Private Sub cmdEmpAttendance_Click()
Call fConn
SQL "select * from tblemployee order by employeeid asc"
If Not RS.EOF Then
'frmEmpAttendance.Show 1
Me.Label1.Caption = "Employee Attendance"
Call empAttendance
Call loadRecords2

Me.ListView4.Visible = True
Me.ListView5.Visible = True
Me.ListView2.Visible = False
'Me.Image4.Visible = False
Else
MsgBox "There are no employees record to transact.", vbExclamation, "Error"
End If

End Sub

Private Sub cmdEmployee_Click()
Me.Label1.Caption = "Employees"
Me.ListView2.Visible = True
Me.ListView4.Visible = False
Me.ListView5.Visible = False
Call loadRecords
Call empAdd
End Sub

Private Sub cmdEmpSetting_Click()
frmEmpSetting.Show 1
End Sub

Private Sub cmdRefresh_Click()
Call loadRecords
End Sub

Private Sub Command2_Click()
Call fConn
SQL "select * from tblemployee order by employeeid asc"
If Not RS.EOF Then
frmTransact.Show 1
Else
MsgBox "There are no employees record to transact.", vbExclamation, "Error"
End If
End Sub

Private Sub Command4_Click()
frmReport.Show 1
End Sub

Private Sub Form_Load()
'Call loadRecords
Label5.Caption = Format$(Now, "dddd  mmmm dd,yyyy")
End Sub

Private Sub ListView1_Click()
On Error GoTo errtrap
If Me.Label1.Caption = "Employees" Then
If Me.ListView1.SelectedItem = Me.ListView1.ListItems(1) Then

frmEmp_Add.Show 1
End If
'If Me.ListView1.SelectedItem = Me.ListView1.ListItems(2) Then
'ok
If Not Me.ListView2.ListItems.Count = 0 And Me.ListView1.SelectedItem = Me.ListView1.ListItems(2) Then

frmEmp_Add.cmdEmp_Add.Caption = "Update"
frmEmp_Add.Command1.Caption = "Update and Close"
Call uploadRecordEdit
frmEmp_Add.Show 1
End If
If Not Me.ListView2.ListItems.Count = 0 And Me.ListView1.SelectedItem = Me.ListView1.ListItems(3) Then
Call deleteRecord
Call loadRecords
End If
ElseIf Me.Label1.Caption = "Employee Attendance" Then
If Me.ListView1.SelectedItem = Me.ListView1.ListItems(1) Then
frmEmpAttendance.Show 1
End If
If Me.ListView1.SelectedItem = Me.ListView1.ListItems(2) Then
If frmMenu.ListView5.ListItems.Count = 0 Then
    MsgBox "There are no records to modify!", vbExclamation, "Error"
    GoTo errtrap
End If
frmEmpAttendance.cmdSave.Caption = "Update"
Call uploadWorkHoursEdit
frmEmpAttendance.Show 1
End If
If Me.ListView1.SelectedItem = Me.ListView1.ListItems(3) Then
Call deleteWorkedHours
Call loadAttendance
End If
End If
errtrap:
End Sub

Private Sub ListView4_Click()
Call loadAttendance
End Sub

Private Sub ListView4_DblClick()
frmEmpAttendance.Show 1
End Sub

Function deleteWorkedHours()
On Error GoTo errtrap
If frmMenu.ListView5.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
     Conn.Execute "DELETE FROM tblemp_attendance WHERE attendanceid = " & Val(frmMenu.ListView5.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  End If
errtrap:
End Function
Private Sub txtSearch_Change()
If Label1.Caption = "Employees" Then
frmMenu.ListView2.ListItems.Clear
'frmMenu.ListView2.HideColumnHeaders = False
  Call fConn
    SQL "Select * from tblemployee where (((lastname) Like '" & txtSearch.Text & "%')) or (((firstname) Like '" & txtSearch.Text & "%')) or (((middlename) Like '" & txtSearch.Text & "%')) or employeeid = " & Val(txtSearch.Text) & ""
    'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
          With frmMenu.ListView2.ListItems
            Set Item = .Add(, , RS!employeeid)
              Item.SubItems(1) = RS!lastname
              Item.SubItems(2) = RS!firstname
              Item.SubItems(3) = RS!middlename
              Item.SubItems(4) = RS!gender
          End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
  Conn.Close
  Set Conn = Nothing
ElseIf Label1.Caption = "Employee Attendance" Then
End If
End Sub
