Public Class frmEditor
#Region "DELCARATIONS"
    Private SQL As New SQLControl : Private Engine As New SQLControl

    Private _MFR As String
    Private _cbx As ComboBox
    Private cbxIndex As Integer
    Private queryHelper As String = Nothing

    Private QueryType As QueryCommand
    Private Enum QueryCommand
        None
        INSERT
        UPDATE
    End Enum
#End Region

    Private Sub frmEditor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If SQL.HasConnection Then InitializeForm() Else MsgBox("No Connection")
    End Sub

#Region "UTILITY SUBS"
    Private Sub InitializeForm()
        Setup_DGV(dgvData)
        'dgvData.Font = New Font("Arial", 9) : dgvData.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        cbxFilter.SelectedIndex = 0 : RefreshForm() : dgvData.Focus()
    End Sub

    Private Sub RefreshForm()
        GetMFR() : MFR_to_DGV()
    End Sub
    Private Sub PartialRefresh()
        Fill_DGV() : PopulateCBX()
    End Sub

    Private Sub GetMFR()
        _MFR = tcMFR.SelectedTab.Text.ToString
        If tcMFR.SelectedTab Is tabGuascor Then _cbx = cbxGuDataSheet : queryHelper = ALL_GUASCOR
        If tcMFR.SelectedTab Is tabMan Then _cbx = cbxMaDataSheet : queryHelper = ALL_MAN
        If tcMFR.SelectedTab Is tabMtu Then _cbx = cbxMtDataSheet : queryHelper = ALL_MTU
    End Sub

    Private Sub MFR_to_DGV()
        SQL.AddParam("@mfr", _MFR)
        SQL.ExecQuery(String.Format("SELECT {0} FROM Engines WHERE mfr=@mfr ORDER BY id", queryHelper))
        PartialRefresh()
    End Sub

    Private Sub Fill_DGV()
        If String.IsNullOrEmpty(SQL.Exception) Then dgvData.DataSource = SQL.DBDS.Tables(0) Else MsgBox(SQL.Exception)
        ColorDVG(dgvData)
        If SQL.RecordCount > 0 Then dgvData.Columns(0).Frozen = True
        If SQL.RecordCount = 0 Then ClearText()
        lblRecords.Text = SQL.RecordCount
    End Sub

    Private Sub PopulateCBX()
        _cbx.Items.Clear()
        For i As Integer = 0 To dgvData.Rows.Count - 1
            If dgvData.Rows(i).Cells(0).Value.ToString IsNot Nothing Then
                _cbx.Items.Add(dgvData.Rows(i).Cells(0).Value.ToString)
            End If
        Next
        Try
            If _cbx.Items.Count > 0 AndAlso dgvData.CurrentRow.Index > -1 Then _cbx.SelectedIndex = dgvData.CurrentRow.Index Else _cbx.SelectedIndex = 0
        Catch ex As Exception : End Try ' DO NOTHING
    End Sub

    Private Sub GetEngine(ID As String) ' POPULATES ALL TEXT BOXES ACCORDING TO SELECTED ENGINE
        Engine.AddParam("@id", ID)
        Engine.ExecQuery("SELECT * FROM Engines WHERE id=@id")
        ' POPULATE FORM BY COLUMN/TAG MATCH
        If String.IsNullOrEmpty(Engine.Exception) AndAlso Engine.RecordCount > 0 Then
            Dim r As DataRow = Engine.DBDS.Tables(0).Rows(0)
            ' LOOP DB COLUMNS & TEXTBOX TAGS FOR MATCHES
            For Each dc As DataColumn In r.Table.Columns
                For Each c As Control In tcMFR.SelectedTab.Controls ' for each control in the selected tab name
                    If Not String.IsNullOrEmpty(c.Tag) AndAlso c.Tag = dc.ColumnName Then
                        c.Text = r(dc).ToString
                        Exit For ' STOP SCANNING AND PROCEED TO NEXT RECORD
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub frmData_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then btnCancelEdit.PerformClick()
        If e.Control AndAlso e.KeyCode = Keys.N Then btnNew.PerformClick()
        If e.Control AndAlso e.KeyCode = Keys.E Then EditMode()
        If e.Control AndAlso e.KeyCode = Keys.S Then btnSave.PerformClick()
        If e.KeyCode = Keys.F5 Then ManualQuery()
    End Sub

    Private Sub ManualQuery()
        Dim q As String = txtManual.Text
        If q.ToLower.Contains("update") Then : SQL.ExecQuery(q) : RefreshForm()
        ElseIf q.ToLower.Contains("insert") Then : SQL.ExecQuery(q) : RefreshForm()
        ElseIf q.ToLower.Contains("select") Then : SQL.ExecQuery(q) : PartialRefresh()
        ElseIf q.ToLower.Contains("drop") Then : MsgBox("You cannot drop tables from this app!")
        ElseIf q = "" Then : RefreshForm() : End If
    End Sub
#End Region

#Region "INDEX CHANGES"
    Private Sub tcMFR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcMFR.SelectedIndexChanged
        If Not String.IsNullOrEmpty(txtSearch.Text) Then WildCard(txtSearch.Text, cbxFilter.Text) Else RefreshForm()
    End Sub

    Private Sub dgvData_Sorted(sender As Object, e As System.EventArgs) Handles dgvData.Sorted
        PopulateCBX()
        ColorDVG(dgvData)
    End Sub

    Private Sub _cbx_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxGuDataSheet.SelectedIndexChanged, cbxMaDataSheet.SelectedIndexChanged, cbxMtDataSheet.SelectedIndexChanged
        Dim i As Integer = _cbx.SelectedIndex
        If i > -1 Then
            ManageIndexChange("cbx")
            GetEngine(_cbx.Text)
        End If
    End Sub
    Private Sub dgvData_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvData.SelectionChanged
        ManageIndexChange("dgv")
    End Sub
    Private Sub ManageIndexChange(sender As String)
        Try
            Dim i As Integer
            Select Case sender
                Case "cbx"
                    i = _cbx.SelectedIndex
                    dgvData.CurrentCell = Nothing
                    dgvData.Rows(i).Selected = True
                    _cbx.SelectedIndex = i
                    dgvData.CurrentCell = dgvData(0, i)
                    _cbx.SelectedIndex = i
                Case "dgv"
                    i = dgvData.CurrentRow.Index
                    _cbx.SelectedIndex = i
            End Select
        Catch ex As Exception : End Try ' DO NOTHING
    End Sub
#End Region

#Region "HOUSEKEEPING"
    Private Sub ToggleLock(Locked As Boolean)
        Dim prefix As String = Nothing
        Select Case _MFR
            Case "Guascor" : prefix = "txtGu"
            Case "MAN" : prefix = "txtMa"
            Case "MTU" : prefix = "txtMt"
        End Select
        For Each c As Control In tcMFR.SelectedTab.Controls
            If c.Name.StartsWith(prefix) Then DirectCast(c, TextBox).ReadOnly = Locked
            If c.Name.EndsWith("Fuel") Or c.Name.EndsWith("Burn") Then
                If Locked Then c.Enabled = False Else c.Enabled = True
            End If
        Next
    End Sub

    Private Sub ClearText()
        _cbx.ResetText()
        For Each c As Control In tcMFR.SelectedTab.Controls
            If TypeOf c Is TextBox Then c.ResetText()
            If c.Name.EndsWith("Fuel") Or c.Name.EndsWith("Burn") Then c.ResetText()
        Next
    End Sub
#End Region

#Region "TOOLBOX"
#Region "Engine Controls"
    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        ' ====  TOGGLE THIS LINE OF CODE TO TRANSFER DATA FROM ACCESS  ====
        'If MsgBox("Are you ready to begin inserting a whole bunch of data?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then : INSERT_ALL() : RefreshForm() : Exit Sub : End If
        QueryType = QueryCommand.INSERT : EditMode()
    End Sub
    Private Sub btnEdit_Click(sender As System.Object, e As System.EventArgs) Handles btnEdit.Click
        QueryType = QueryCommand.UPDATE : EditMode()
    End Sub
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim i As Integer = dgvData.CurrentRow.Index
        Dim id As String = dgvData.Rows(i).Cells(0).Value
        If MsgBox(String.Format("Are you sure you want to delete {0} forever?", id), MsgBoxStyle.YesNo, "CONFIRM") = MsgBoxResult.Yes Then
            SQL.AddParam("@id", id)
            SQL.ExecQuery("DELETE FROM Engines WHERE id=@id")
            If String.IsNullOrEmpty(SQL.Exception) Then MsgBox(String.Format("Sucessfully deleted {0}", id)) Else MsgBox(SQL.Exception)
        End If
        RefreshForm()
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If MsgBox("Do the changes you make look good enough to save?", MsgBoxStyle.YesNo, "CONFIRM") = MsgBoxResult.Yes Then SQLCommand() : ViewMode()
    End Sub
    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        ClearText()
    End Sub
    Private Sub btnCancelEdit_Click(sender As System.Object, e As System.EventArgs) Handles btnCancelEdit.Click
        ViewMode()
    End Sub
#End Region
#Region "SQL Modes"
    Private Sub EditMode() ' EDIT MODE
        cbxIndex = _cbx.SelectedIndex
        ' CALC ELE POWER
        btnCancelEdit.Enabled = True
        btnSave.Enabled = True
        btnClear.Enabled = True
        dgvData.Enabled = False
        btnNew.Enabled = False
        btnDelete.Enabled = False
        btnEdit.Enabled = False
        cbxFilter.Enabled = False
        txtSearch.Enabled = False
        ToggleLock(False)
        _cbx.DropDownStyle = ComboBoxStyle.Simple
        Select Case QueryType
            Case QueryCommand.INSERT
                lblMode.Text = "New Engine Mode" : pnlMode.BackColor = Color.LightGreen : dgvData.GridColor = Color.Green
                ClearText()
            Case QueryCommand.UPDATE
                lblMode.Text = "Edit Mode" : pnlMode.BackColor = Color.Tomato : dgvData.GridColor = Color.Tomato
        End Select
    End Sub
    Private Sub ViewMode() ' VIEW MODE
        QueryType = QueryCommand.None
        lblMode.Text = "View Mode" : pnlMode.BackColor = Color.Transparent : dgvData.GridColor = Color.White
        btnCancelEdit.Enabled = False
        btnSave.Enabled = False
        btnClear.Enabled = False
        dgvData.Enabled = True
        btnNew.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        cbxFilter.Enabled = True
        txtSearch.Enabled = True
        _cbx.Enabled = True
        _cbx.DropDownStyle = ComboBoxStyle.DropDown
        ToggleLock(True)
        RefreshForm()
        _cbx.Focus() : _cbx.SelectedIndex = cbxIndex
    End Sub
    Private Sub SQLCommand() ' SQL COMMAND (INSERT & UPDATE)
        Select Case QueryType
            Case QueryCommand.INSERT
                Dim cols As String = ""
                Dim params As String = ""
                For Each k As KeyValuePair(Of String, String) In _GetParams(tcMFR.SelectedTab)
                    cols = (cols & k.Key & ",")
                    params = (params & "@" & k.Key & ",")
                    SQL.AddParam("@" & k.Key, k.Value)
                Next
                SQL.ExecQuery(String.Format("INSERT INTO Engines ({0}) VALUES ({1})", cols.Remove(cols.Length - 1, 1), params.Remove(params.Length - 1, 1)))
            Case QueryCommand.UPDATE
                Dim command As String = ""
                For Each k As KeyValuePair(Of String, String) In _GetParams(tcMFR.SelectedTab)
                    command &= (k.Key & "=@" & k.Key & ",")
                    SQL.AddParam("@" & k.Key, k.Value)
                Next
                SQL.ExecQuery(String.Format("UPDATE Engines SET {0} WHERE id=@id", command.Remove(command.Length - 1, 1)))
        End Select
        If Not String.IsNullOrEmpty(SQL.Exception) Then MsgBox(SQL.Exception)
    End Sub
#End Region
#Region "Other Tools"
    ' FILTER DATA
    Private Sub tsSearch_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSearch.TextChanged
        If txtSearch.Text <> "" Then WildCard(txtSearch.Text, cbxFilter.Text) Else RefreshForm()
    End Sub
    Public Sub WildCard(Key As String, Filter As String)
        GetMFR()
        SQL.AddParam("@key", String.Format("%{0}%", Key)) : SQL.AddParam("@mfr", _MFR)
        SQL.ExecQuery(String.Format("SELECT {0} FROM Engines WHERE mfr=@mfr AND {1} LIKE @key ORDER BY id", queryHelper, Filter))
        PartialRefresh()
    End Sub
    ' SIZE TOOLS
    Private Sub btnWindow_Click(sender As System.Object, e As System.EventArgs) Handles btnWindow.Click
        If Me.WindowState = System.Windows.Forms.FormWindowState.Maximized Then Me.WindowState = FormWindowState.Normal
        Me.Width = 1024 : Me.Height = 740 : dgvData.Location = New Point(643, 85) : Default_DGV(dgvData.Left)
        If chkExpand.Checked Then chkExpand.Checked = False
        txtSearch.ResetText()
    End Sub
    Private Sub chkExpand_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkExpand.CheckedChanged
        If chkExpand.Checked Then
            chkExpand.Image = My.Resources.Shrink : tcMFR.Hide() : dgvData.Left = 0 : Default_DGV()
        Else
            chkExpand.Image = My.Resources.Expand : dgvData.Location = New Point(643, 85) : Default_DGV(dgvData.Left) : tcMFR.Show()
        End If
    End Sub
    Private Sub Default_DGV(Optional ShrinkSize As Integer = 0)
        dgvData.Width = Me.Width - (ShrinkSize) - 16 : dgvData.Height = Me.Height - dgvData.Top - 95
    End Sub
#End Region
#End Region
End Class