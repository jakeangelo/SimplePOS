Attribute VB_Name = "Connection"
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim command As New ADODB.command

Public Sub connect()
    If conn.State <> 0 Then conn.Close
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + App.Path + "\db\posdb.accdb;Persist Security Info=False"
End Sub
