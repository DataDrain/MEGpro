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

    ' DGV SETUP / UTLITY
    Public Sub Setup_DGV(dgv As DataGridView)
        dgv.Font = New Font("Arial", 9) : dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
    End Sub
    Public Sub ColorDVG(dgv As DataGridView)
        For i As Integer = 0 To dgv.Rows.Count - 1
            dgv.Rows(i).Cells(0).Style.BackColor = Color.LightGray
        Next
    End Sub
    Public Function _get(ds As DataSet, colName As String, Optional row As Integer = 0) As Object
        Return ds.Tables(0).Rows(row)(colName)
    End Function

    ' MULTIPLE TAB OBJECT BEHAVIOR
    Public Sub UpdateObj(obj As Object, status As String, color As Color, Optional backcolor As Color = Nothing)
        obj.Text = status : obj.ForeColor = color : DirectCast(obj, Label).BackColor = backcolor
    End Sub

    Public Function isNullorNumeric(tBox As TextBox) As Boolean
        If String.IsNullOrEmpty(tBox.Text) Or IsNumeric(tBox.Text) Then Return True Else Return False
    End Function

    ' HEAT TAB OBJECT BEHAVIOR
    Public Function isNotNullorInvalid(ByVal tBox As TextBox) As Boolean
        If Not String.IsNullOrEmpty(tBox.Text.ToString) AndAlso IsNumeric(tBox.Text.ToString) Then Return True Else Return False
    End Function

    Public Function withinRange(ByVal tBox As TextBox) As Boolean
        If tBox.Text < 30 Or tBox.Text > 50 Then Return False Else Return True
    End Function

    Public Sub ToggleHeatControls(ByVal c1 As Control, ByVal c2 As Control, ByVal Toggle As Boolean, ByVal Checked As Boolean, ByVal Enabled As Boolean, Optional ByVal ResetText As Boolean = False)
        If Checked Then : DirectCast(c1, RadioButton).Checked = Toggle : DirectCast(c2, RadioButton).Checked = Toggle : End If
        If Enabled Then : c1.Enabled = Toggle : c2.Enabled = Toggle : End If
        If ResetText Then c1.Text = 0 : c2.Text = 0
    End Sub

    Public Sub ViewMode(Multiple As String, tab As TabPage)
        For Each c As Control In tab.Controls
            If TypeOf c Is Label Then
                Select Case Multiple
                    Case "20"
                        If c.Name.ToString.EndsWith("80") Or c.Name.ToString.EndsWith("60") Or c.Name.ToString.EndsWith("40") Then c.Visible = True
                        If c.Name.ToString.EndsWith("75") Or c.Name.ToString.EndsWith("50") Then c.Visible = False
                    Case "25"
                        If c.Name.ToString.EndsWith("75") Or c.Name.ToString.EndsWith("50") Then c.Visible = True
                        If c.Name.ToString.EndsWith("80") Or c.Name.ToString.EndsWith("60") Or c.Name.ToString.EndsWith("40") Then c.Visible = False
                End Select
            End If
        Next
    End Sub

    ' FORMATTING
    Public Function NoDec(value As Object) As Object
        Return String.Format("{0:n0}", value)
    End Function
    Public Function OneDec(value As Object) As Object
        Return String.Format("{0:n1}", value)
    End Function

    ' MAIN FORM NAVIGATION
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
                    Case 0 : If count <> 0 Then : tc.SelectedIndex += 1 : frmMain.btnBack.Enabled = True : Else : MsgBox("You have 0 results found, you need at least 1 to continue.") : End If
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