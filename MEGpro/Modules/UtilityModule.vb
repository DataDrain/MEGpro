Module UtilityModule
    ' AGNOSTIC PARAMETER DICTIONARY
    Public Function _GetParams(Parent As Object) As Dictionary(Of String, String)
        Dim ParamDict As New Dictionary(Of String, String)
        For Each c As Control In Parent.controls
            If Not String.IsNullOrEmpty(c.Tag) AndAlso Not String.IsNullOrEmpty(c.Text) Then
                If Not ParamDict.ContainsKey(c.Tag) Then ParamDict.Add(c.Tag, c.Text)
            End If
        Next
        Return ParamDict
    End Function

    Public Function isTxtValid(tBox As TextBox) As Boolean
        If String.IsNullOrEmpty(tBox.Text.ToString) Then Return True
        If IsNumeric(tBox.Text.ToString) Then Return True Else MsgBox(String.Format("{0} is non-numeric text.", tBox.Text)) : Return False
    End Function
End Module
