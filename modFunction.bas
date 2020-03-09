Attribute VB_Name = "modFunctionAdd"
Function empAdd()
Dim ITM As ListItem
'Dim LSI As ListSubItem

'Set ListView3.ColumnHeaderIcons = ImageList1
'ListView3.ColumnHeaders.Add , , "Item", , , "b"
'frmMenu.ListView1.ListItems.Clear
Set frmMenu.ListView1.SmallIcons = frmMenu.ImageList1
With frmMenu.ListView1.ListItems
.Clear
Set ITM = .Add(, , "Add Employee", , "add")
Set ITM = .Add(, , "Edit Employee", , "edit")
Set ITM = .Add(, , "Delete Employee", , "delete")
Set ITM = .Add(, , "Search Employee", , "search")
'Set LSI = ITM.ListSubItems.Add(, , "Subitem " & 1 & 2)
End With
End Function
Function empAttendance()
Dim ITM As ListItem

Set frmMenu.ListView1.SmallIcons = frmMenu.ImageList2
With frmMenu.ListView1.ListItems
.Clear
Set ITM = .Add(, , "Add Registration", , "add")
Set ITM = .Add(, , "Edit Registration", , "edit")
Set ITM = .Add(, , "Delete Registration", , "delete")
End With
End Function
Function loadRecords()
frmMenu.ListView2.ListItems.Clear
frmMenu.ListView2.HideColumnHeaders = False
  Call fConn
    SQL "Select * from tblemployee order by employeeid asc"
    'RS.Open SQL, Conn, adOpenDynamic
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
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function
Function loadRecords2()
frmMenu.ListView4.ListItems.Clear
'frmMenu.ListView4.HideColumnHeaders = False
  Call fConn
    SQL "Select * from tblemployee order by employeeid asc"
    'RS.Open SQL, Conn, adOpenDynamic
      'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
          With frmMenu.ListView4.ListItems
            Set Item = .Add(, , RS!employeeid)
              Item.SubItems(1) = RS!lastname
              Item.SubItems(2) = RS!firstname
              Item.SubItems(3) = RS!middlename
              'Item.SubItems(4) = RS!gender
          End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function
Function addEmployee()
Dim Increment As Integer
On Error GoTo errtrap
If frmEmp_Add.cmdEmp_Add.Caption = "Save" Then
Call fConn
'SQL "select (max(employeeID)+1) as incremented from tblemployee;"
'Increment = RS!incremented
    'SQL "select * from tblemployee;"
'RS.MoveLast
'If Not RS.EOF Then
    'RS.MoveFirst
    'Do While Not RS.EOF
   ' i = RS!employeeid
    'RS.MoveNext
    'Loop
'End If
'RS.Close
'Increment = i + 1
Increment = Val(frmEmp_Add.Label2.Caption)
With frmEmp_Add
SQL "Insert into tblemployee values(" & Increment & ",'" & LCase(.txtLastName) & "','" & LCase(.txtFirstName) & "','" & LCase(.txtMiddleName) & "','" & LCase(.cboGender) & "','" & LCase(.txtTIN) & "','" & LCase(.txtStreet) & "','" & LCase(.txtProvince) & "','" & LCase(.txtRegion) & "','" & LCase(.txtPhone) & "','" & LCase(.txtEmail) & "','" & LCase(.txtNotes) & "','" & LCase(.cboCivilStatus) & "'," & Val(.txtDependent) & "," & Val(.cboGrade) & "," & Val(.cboStep) & "," & Val(Trim(Left(.cboStation, 3))) & "," & Val(Trim(Left(.cboDivision, 3))) & "," & Val(Trim(Left(.cboRegion, 3))) & "," & Val(Trim(Left(.cboPosition, 3))) & ",'" & .DTPicker1 & "','" & .DTPicker2 & "')"
End With
'SQL "Insert into tblemployee values(" & Increment & ",'" & frmEmp_Add.txtLastName & "','" & frmEmp_Add.txtFirstName & "','" & frmEmp_Add.txtMiddleName & "','" & frmEmp_Add.cboGender & "','" & frmEmp_Add.txtTIN & "','" & frmEmp_Add.txtStreet & "','" & frmEmp_Add.txtProvince & "','" & frmEmp_Add.txtRegion & "','" & frmEmp_Add.txtPhone & "','" & frmEmp_Add.txtEmail & "','" & frmEmp_Add.txtNotes & "','" & frmEmp_Add.cboCivilStatus & "'," & Val(frmEmp_Add.txtDependent) & "," & Val(frmEmp_Add.cboGrade) & "," & Val(frmEmp_Add.cboStep) & "," & Val(frmEmp_Add.cboStation) & "," & Val(frmEmp_Add.cboDivision) & "," & Val(frmEmp_Add.cboRegion) & "," & Val(frmEmp_Add.cboPosition) & ",'" & frmEmp_Add.DTPicker1 & "','" & frmEmp_Add.DTPicker2 & "')"
MsgBox "Data saved."
Conn.Close
Set Conn = Nothing
Else
Call editEmployee
End If
errtrap:
 Select Case Err.Number
    Case -2147467259
      MsgBox "The name already exists in the database", vbCritical, "Error"
    Case -2147217887
      MsgBox "Please fill up all the required fields.", vbCritical, "Error"
 '   Case Else
 '     MsgBox Err.Description, vbCritical, "The system encountered an error"
  End Select
End Function
Function loadEmpID()
frmEmp_Add.txtLastName.SetFocus
Call fConn
SQL "select * from tblemployee order by employeeid asc"
If Not RS.EOF Then
    RS.MoveFirst
    Do While Not RS.EOF
    frmEmp_Add.Label2.Caption = Val(RS!employeeid) + 1
    RS.MoveNext
    Loop
    Conn.Close
Else
frmEmp_Add.Label2.Caption = "1"
End If
End Function
Function clearField()
With frmEmp_Add
.txtLastName = ""
.txtFirstName = ""
.txtMiddleName = ""
'.cboGender = ""
.txtTIN = ""
.txtStreet = ""
.txtProvince = ""
.txtRegion = ""
.txtPhone = ""
.txtEmail = ""
.txtNotes = ""
'.cboCivilStatus = ""
.txtDependent = ""
'.cboGrade = ""
'.cboStep = ""
'.cboStation = ""
'.cboDivision = ""
'.cboRegion = ""
'.cboPosition = ""
End With
End Function

Function editEmployee()

If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Call fConn
    'SQL = "UPDATE tblEmployee SET lastName = '" & .txtLastName & "', firstame = '" & Me.txtCatAddPercent.Text & "' " & _
     '     "WHERE (((tblCategory.categoryID)=" & Val(frmCategory.ListView.SelectedItem.Text) & "));"
    'Call dbConnect
     ' Conn.Execute SQL
     With frmEmp_Add
        SQL "UPDATE tblEmployee SET lastname =  '" & .txtLastName & "',firstname = '" & .txtFirstName & "',middleName = '" & .txtMiddleName & "',gender = '" & LCase(.cboGender) & "',TIN = '" & .txtTIN & "',street_Brngy = '" & .txtStreet & "',province_city = '" & .txtProvince & "',region_emp = '" & .txtRegion & "',cell_phone = '" & .txtPhone & "',email = '" & .txtEmail & "', notes = '" & .txtNotes & "',civilstatus = '" & .cboCivilStatus & "',dependent = " & Val(.txtDependent) & ",grade = " & Val(.cboGrade) & ",step = " & Val(.cboStep) & " ,station = " & Val(.cboStation) & ",division = " & Val(.cboDivision) & ",region = " & Val(.cboRegion) & ",dateOfBirth = '" & .DTPicker1 & "',employmentDate = '" & .DTPicker2 & "' " & _
        "WHERE (((employeeid)= " & Val(frmMenu.ListView2.SelectedItem.Text) & "))"
     End With
    Conn.Close
    Set Conn = Nothing
    'Unload Me
  Else
    Cancel = True
  End If
  'Exit Sub
'errtrap:
End Function
Function uploadRecordEdit()
On Error GoTo errtrap
Call fConn
SQL "Select * from tblemployee WHERE employeeid = " & Val(frmMenu.ListView2.SelectedItem.Text) & " "
With frmEmp_Add
.Label2.Caption = RS!employeeid
.txtLastName = RS!lastname
.txtFirstName = RS!firstname
.txtMiddleName = RS!middlename
.cboGender = RS!gender
.txtTIN = RS!tin
.txtStreet = RS!street_brngy
.txtProvince = RS!province_city
.txtRegion = RS!region_emp
.txtPhone = RS!cell_phone
.txtEmail = RS!email
.txtNotes = RS!notes
.cboCivilStatus = RS!civilstatus
.txtDependent = RS!dependent
.cboGrade = RS!grade
.cboStep = RS!step
.cboPosition = RS!t_position
.cboStation = RS!station
.cboDivision = RS!division
.cboRegion = RS!region

.DTPicker1 = RS!dateofbirth
.DTPicker2 = RS!employmentdate
End With
Conn.Close
errtrap:
End Function
Function deleteRecord()
'On Error GoTo errtrap
If frmMenu.ListView2.ListItems.Count = 0 Then
    MsgBox "There are no records to modify or delete!", vbExclamation, "Error"
    GoTo errtrap
  End If
  If MsgBox("Are you sure you want to delete the selected record?", vbYesNo, "Delete") = vbYes Then
    Call fConn
    'MsgBox (frmMenu.ListView2.SelectedItem.Text)
     Conn.Execute "DELETE FROM tblEmployee WHERE employeeid = " & Val(frmMenu.ListView2.SelectedItem.Text) & ";"
    Conn.Close
    Set Conn = Nothing
    'frmMenu.ListView2.ListItems.Remove (frmMenu.ListView2.SelectedItem.Index)
    MsgBox "The selected record was deleted", vbExclamation, "Delete"
  Else
    Cancel = True
  End If
errtrap:
End Function
Function uploadWorkHoursEdit()
On Error GoTo errtrap
Call fConn
SQL "Select * from tblemp_attendance WHERE attendanceid = " & Val(frmMenu.ListView5.SelectedItem.Text) & " "
With frmEmpAttendance
.DTPicker1 = RS!datestarted
.DTPicker2 = RS!dateended
.txtWorkedHours = RS!workedhours
.txtAbs_Tar = RS!absent_tardy
End With
'Conn.Close
errtrap:
End Function







