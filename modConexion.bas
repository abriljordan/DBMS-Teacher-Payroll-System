Attribute VB_Name = "modConexion"
Option Explicit
Public Conn As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Item As ListItem
Function SQL(sqlstring As String)
    If RS.State = adStateOpen Then RS.Close
    RS.Open sqlstring, Conn, adOpenDynamic, adLockOptimistic
End Function
Function SQL2(sqlstring As String)
    If RS2.State = adStateOpen Then RS2.Close
    RS2.Open sqlstring, Conn, adOpenDynamic, adLockOptimistic
End Function
Function fConn()
    Set Conn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    Conn.ConnectionString = "ODBC;Driver={PostgreSQL};DSN=PostgreSQL;UID=;PWD="
    'Conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL"
    Conn.Open
    RS.ActiveConnection = Conn
End Function

'***************************************************

Public Function DigitOnly(pintKeyAscii As Integer) As Integer
    If (Chr$(pintKeyAscii) >= "0" And Chr$(pintKeyAscii) <= "9") _
    Or (pintKeyAscii < 32) Then
        DigitOnly = pintKeyAscii
    Else
        DigitOnly = 0
    End If
End Function

