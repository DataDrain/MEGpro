Module AccessToSQL
    ' THIS MODULE IS SOLELY USED TO TRANSFER DATA FROM ACCESS TO SQLce
    Private SQL As New SQLControl
    Private Rows As New List(Of String)
    Private TableName As String = "MAN"
    Private Values As String '=

    Public Sub INSERT_ALL()
        Dim Query As String = Nothing
        BuildList()
        For Each s As String In Rows
            Try
                Query = String.Format("INSERT INTO {0} ({1}) VALUES ({2})", TableName, Values, s)
                SQL.ExecQuery(Query)
                If Not String.IsNullOrEmpty(SQL.Exception) Then MsgBox(SQL.Exception)
            Catch ex As Exception : End Try ' DO NOTHING
        Next
        MsgBox(String.Format("Done Inserting"))
    End Sub
    Private Sub BuildList()
        ' ADD ALL ROWS.ADD(" 'id', 'mfr', 'model', etc... ") CODE HERE
    End Sub
End Module