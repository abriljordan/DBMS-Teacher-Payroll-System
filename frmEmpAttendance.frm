VERSION 5.00
Begin VB.Form frmEmpAttendance 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
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
      Left            =   6240
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Caption         =   "Worked Hour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   7215
      Begin VB.Frame Frame6 
         BackColor       =   &H80000003&
         Caption         =   "Total Worked Hour "
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
         TabIndex        =   20
         Top             =   1680
         Width           =   2775
         Begin VB.Label Label5 
            BackColor       =   &H80000003&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.TextBox txtAbs_Tar 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000003&
         Height          =   615
         Left            =   3480
         TabIndex        =   15
         Top             =   1440
         Width           =   3615
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000003&
            Caption         =   "Day(s)"
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
            Left            =   2040
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000003&
            Caption         =   "Hour(s)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000003&
         Caption         =   "Absent/Tardy"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000003&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6975
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000003&
            Caption         =   "Hour(s)"
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
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000003&
            Caption         =   "Day(s)"
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
            Left            =   3000
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000003&
            Caption         =   "Week(s)"
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
            Left            =   5280
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtWorkedHours 
         Height          =   405
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Hour(s):"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   2280
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Hour(s):"
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
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H80000003&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         Begin VB.PictureBox DTPicker2 
            Height          =   375
            Left            =   4320
            ScaleHeight     =   315
            ScaleWidth      =   2475
            TabIndex        =   5
            Top             =   240
            Width           =   2535
         End
         Begin VB.PictureBox DTPicker1 
            Height          =   375
            Left            =   1320
            ScaleHeight     =   315
            ScaleWidth      =   2475
            TabIndex        =   2
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000003&
            Caption         =   "to"
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
            Left            =   3960
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Date: "
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
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmEmpAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private calWorkedHours As Double
Private calAbsTarHours As Double


Private Sub Check1_Click()
If Check1.Value = False Then
Me.Frame5.Visible = False
Me.Label4.Visible = False
Me.txtAbs_Tar.Visible = False
Me.txtAbs_Tar.Text = ""
Else
Me.Frame5.Visible = True
Me.Label4.Visible = True
Me.txtAbs_Tar.Visible = True
End If
End Sub

Function addWorkedHours()
Dim Increment As Integer
If Option6.Value = True Then
calAbsTarHours = Val(txtAbs_Tar) * 8
ElseIf Option5.Value = True Then
calAbsTarHours = Val(txtAbs_Tar.Text)
End If
Call fConn
SQL "Select * from tblemp_attendance;"
If Not RS.EOF Then
SQL "select (max(attendanceid)+1) as incremented from tblemp_attendance;"
Increment = RS!incremented
Else
Increment = 1
End If
With frmEmpAttendance
SQL "Insert into tblemp_attendance values(" & Increment & ",'" & .DTPicker1 & "','" & .DTPicker2 & "'," & Val(.Label5) & "," & calAbsTarHours & "," & Val(frmMenu.ListView4.SelectedItem) & ")"
End With
MsgBox "Data saved."
Conn.Close
Set Conn = Nothing
End Function

Function updateWorkedHours()
'On Error GoTo errtrap
If frmMenu.ListView5.ListItems.Count = 0 Then
    MsgBox "There are no records to modify!", vbExclamation, "Error"
    GoTo errtrap
End If
If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Call fConn
     With frmEmpAttendance
        SQL "UPDATE tblemp_attendance SET datestarted =  '" & .DTPicker1 & "',dateended = '" & .DTPicker2 & "', workedhours = " & Val(.txtWorkedHours) & ", absent_tardy = " & Val(.txtAbs_Tar) & " " & _
        "WHERE (((attendanceid)= " & Val(frmMenu.ListView5.SelectedItem.Text) & "))"
     End With
    Conn.Close
    Set Conn = Nothing
  End If
'errtrap:
End Function

Private Sub cmdSave_Click()
If Me.cmdSave.Caption = "Save" Then
Call addWorkedHours
Call loadAttendance
Unload Me
ElseIf Me.cmdSave.Caption = "Update" Then
Call updateWorkedHours
Call loadAttendance
Unload Me
End If
End Sub

Private Sub Form_Load()

  Call fConn
  SQL "Select * from tblemployee where employeeid = " & Val(frmMenu.ListView4.SelectedItem) & " order by employeeid asc"
  Me.Caption = RS!employeeid & Space(2) & RS!lastname & Space(2) & "," & Space(2) & RS!firstname & Space(2) & RS!middlename
  Conn.Close
  Set Conn = Nothing
Me.Frame5.Visible = False
Me.Label4.Visible = False
Me.txtAbs_Tar.Visible = False
End Sub

Private Sub Option1_Click()
Me.Label3.Caption = "Hour(s):"
Me.Label5.Caption = txtWorkedHours
End Sub

Private Sub Option2_Click()
Me.Label3.Caption = "Day(s):"
calWorkedHours = Val(txtWorkedHours) * 8
Me.Label5.Caption = calWorkedHours
End Sub

Private Sub Option3_Click()
Me.Label3.Caption = "Week(s):"
calWorkedHours = Val(txtWorkedHours) * 40
Me.Label5.Caption = calWorkedHours
End Sub

Private Sub Option4_Click()
Me.Label3.Caption = "Month(s):"
End Sub

Private Sub Option5_Click()
Me.Label4.Caption = "Hour(s):"
End Sub

Private Sub Option6_Click()
Me.Label4.Caption = "Day(s):"
End Sub

Private Sub txtAbs_Tar_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub

Private Sub txtWorkedHours_Change()
If Option1.Value = True Then
Me.Label5.Caption = txtWorkedHours
ElseIf Option2.Value = True Then
calWorkedHours = Val(txtWorkedHours) * 8
Me.Label5.Caption = calWorkedHours
ElseIf Option3.Value = True Then
calWorkedHours = Val(txtWorkedHours) * 40
Me.Label5.Caption = calWorkedHours
End If
End Sub

Private Sub txtWorkedHours_KeyPress(KeyAscii As Integer)
KeyAscii = DigitOnly(KeyAscii)
End Sub
