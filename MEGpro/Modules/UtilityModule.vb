Module UtilityModule
    Private SQL As New SQLControl

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

    ' MAIN FORM UTLITY
    Public Function isNullorNumeric(tBox As TextBox) As Boolean
        If String.IsNullOrEmpty(tBox.Text) Or IsNumeric(tBox.Text) Then Return True Else Return False
    End Function

    Public Function isNotNullorInvalid(ByVal tBox As TextBox) As Boolean
        If Not String.IsNullOrEmpty(tBox.Text.ToString) AndAlso IsNumeric(tBox.Text.ToString) Then Return True Else Return False
    End Function

    Public Sub UpdateObj(obj As Object, status As String, color As Color)
        obj.Text = status : obj.ForeColor = color
    End Sub

    Public Sub Navigate(ButtonPushed As String, tc As TabControl, count As Integer)
        Select Case ButtonPushed
            Case "btnBack"
                Select Case tc.SelectedIndex
                    Case 1 : tc.SelectedIndex -= 1 : frmMain.btnBack.Enabled = False
                    Case 2 : tc.SelectedIndex -= 1
                    Case 3 : tc.SelectedIndex -= 1
                    Case 4 : tc.SelectedIndex -= 1
                    Case 5 : tc.SelectedIndex -= 1 : frmMain.btnNext.Enabled = True
                End Select
            Case "btnNext"
                Select Case tc.SelectedIndex
                    Case 0 : If count <> 0 Then : tc.SelectedIndex += 1 : frmMain.btnBack.Enabled = True : Else : MsgBox("No records found") : End If
                    Case 1 : tc.SelectedIndex += 1
                    Case 2
                        Try
                            tc.SelectedIndex += 1
                        Catch ex As Exception : End Try ' DO NOTHING
                    Case 3 : tc.SelectedIndex += 1
                    Case 4 : tc.SelectedIndex += 1 : frmMain.btnNext.Enabled = False
                    Case 5
                End Select
        End Select
    End Sub
End Module