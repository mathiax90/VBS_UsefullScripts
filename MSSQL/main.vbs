Sub main()
    'con vars
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String

    'Init sql connection
    sConnString = "Provider=SQLOLEDB;Data Source=tcp:127.0.0.1,1433;" & _
                "Initial Catalog=TESTDB;" & _
                "User ID=sa;Password=sa;"
                
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.Open sConnString

    SqlQuery = "select * from test_table"
    rs.Open SqlQuery, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Do
        If rs.EOF Then Exit Do
        'do something
        rs1.MoveNext
    Loop

    rs.Close
    'cleaning
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing

End Sub