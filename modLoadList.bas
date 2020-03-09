Attribute VB_Name = "modLoadList"
Function stationList()
  Call fConn
    SQL "Select * from tblstation order by stationid asc"
      'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
             With frmEmp_Add.cboStation
                .AddItem (RS!stationid & Space(3) & StrConv(RS!station_name, vbProperCase))
             End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function

Function divisionList()
  Call fConn
    SQL "Select * from tbldivision order by divisionid asc"
      If Not RS.EOF Then
        RS.MoveFirst
        Do While Not RS.EOF
             With frmEmp_Add.cboDivision
                .AddItem (RS!divisionid & Space(3) & StrConv(RS!division_name, vbProperCase))
             End With
          RS.MoveNext
          'DoEvents
        Loop
      Else
        MsgBox "No Division records listed."
      End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function

Function regionList()
  Call fConn
    SQL "Select * from tblregion order by regionid asc"
    'RS.Open SQL, Conn, adOpenDynamic
      'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
             With frmEmp_Add.cboRegion
                .AddItem (RS!regionid & Space(3) & StrConv(RS!region_name, vbUpperCase))
             End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function

Function positionList()
  Call fConn
    SQL "Select * from tblposition order by positionid asc"
    'RS.Open SQL, Conn, adOpenDynamic
      'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
             With frmEmp_Add.cboPosition
                .AddItem (RS!positionid & Space(5) & StrConv(RS!position_name, vbProperCase))
             End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function

Function loadAttendance()
Call fConn
frmMenu.ListView5.ListItems.Clear
SQL "Select * from tblemp_attendance where employeeid = " & Val(frmMenu.ListView4.SelectedItem) & " order by attendanceid asc"
'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
          With frmMenu.ListView5.ListItems
            Set Item = .Add(, , RS!attendanceid)
              Item.SubItems(1) = RS!datestarted
              Item.SubItems(2) = RS!dateended
              Item.SubItems(3) = RS!workedhours
              Item.SubItems(4) = RS!absent_tardy
          End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If

Conn.Close
Set Conn = Nothing
End Function
Function deductionTypeList()
  Call fConn
    SQL "Select * from tbltypeofdeduction order by deductionid asc"
    'RS.Open SQL, Conn, adOpenDynamic
      'If Not RS.EOF Then
        'RS.MoveFirst
        Do While Not RS.EOF
             With frmTransact.cboDeductionType
                .AddItem (RS!deductionid & Space(5) & StrConv(RS!deductionname, vbProperCase))
             End With
          RS.MoveNext
          'DoEvents
        Loop
      'End If
    'RS.Close
  Conn.Close
  Set Conn = Nothing
End Function
Function allowanceTypeList()
  Call fConn
    SQL "Select * from tbltypeofallowance order by type_allowanceid asc"
        Do While Not RS.EOF
             With frmTransact.cboAllowance
                .AddItem (RS!type_allowanceid & Space(5) & StrConv(RS!allowance_descrip, vbProperCase))
             End With
          RS.MoveNext
        Loop
  Conn.Close
  Set Conn = Nothing
End Function
