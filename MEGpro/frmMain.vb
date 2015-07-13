Imports System.Drawing
Imports System.Threading

Public Class frmMain
    Public Version As String = "1.5.0"
#Region "DECLARATIONS"
    Private SQL As New SQLControl
    Public PowFactor As Single = 1
    Public Index As Integer

    Private MyGenset As Genset
    Private GensetList As New List(Of Genset)

    Public bmp As Bitmap
#End Region

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        InitializeForm()
    End Sub

#Region "HOUSEKEEPING & UTILITY"
    Private Sub InitializeForm()
        miVersion.Text = String.Format("Version: {0}", Version)
        cbxEngCoolant.SelectedIndex = 0 : cbxPrimaryCir.SelectedIndex = 0 : cbxFilter.SelectedIndex = 0 : FillGensetDGVCols(dgvGensets)
        Setup_DGV(dgvCompare)
        'For Each tp As TabPage In tcMain.TabPages
        '    If tp.Text <> "Choose Application" Then tp.Enabled = False
        'Next
    End Sub

    Private Sub TextBox_Click_SelectAll(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEPmin.Click, txtEPmax.Click, txtMinExTemp.Click, txtSteam.Click, txtFeed.Click, txtPrimaryInlet.Click, txtPrimaryOutlet.Click, txt2ndInlet.Click, txt2ndOutlet.Click
        DirectCast(sender, TextBox).SelectAll()
    End Sub
#End Region

#Region "QUERY SYNTHESIS"
    Private Sub SynthQuery()
        If canQuery() Then
            Dim strMFR As String = getFilters(Filter.mfr)
            Dim strRPM As String = getFilters(Filter.rpm)
            Dim strFuel As String = getFilters(Filter.fuel)
            Dim strBurn As String = getFilters(Filter.burn_type)
            Dim strNOx As String = getFilters(Filter.nox)
            Dim strMin As String = getFilters(Filter.min)
            Dim strMax As String = getFilters(Filter.max)
            Dim strVolts As String = getFilters(Filter.voltage)
            Dim query As String = "SELECT id, mfr, model, rpm, fuel, burn_type, nox, elepow100, voltage FROM Engines WHERE "
            If Not String.IsNullOrEmpty(strMFR) Then query &= strMFR & vbCrLf
            If Not String.IsNullOrEmpty(strRPM) Then query &= "AND " & strRPM & vbCrLf
            If Not String.IsNullOrEmpty(strFuel) Then query &= "AND " & strFuel & vbCrLf
            If Not String.IsNullOrEmpty(strBurn) Then query &= "AND " & strBurn & vbCrLf
            If Not String.IsNullOrEmpty(strNOx) Then query &= "AND " & strNOx & vbCrLf
            If Not String.IsNullOrEmpty(strMin) Then query &= "AND " & strMin & " "
            If Not String.IsNullOrEmpty(strMax) Then query &= "AND " & strMax & vbCrLf
            If Not String.IsNullOrEmpty(strVolts) Then query &= "AND " & strVolts & vbCrLf
            query &= "ORDER BY ElePow100, RPM"
            'MsgBox(query) ' <---- VIEW THE QUERY
            SQL.ExecQuery(query)
            If String.IsNullOrEmpty(SQL.Exception) Then dgvCompare.DataSource = SQL.DBDS.Tables(0) Else MsgBox(SQL.Exception)
            OptionsFound()
        Else
            Try
                UpdateObj(lblRecords, 0, Color.Red, Color.Gainsboro) : SQL.DBDS.Clear()
            Catch ex As Exception : End Try ' DO NOTHING
        End If
    End Sub

    Private Function canQuery() As Boolean
        Dim chksValid As Boolean = False : Dim txtValid As Boolean = False
        For Each c As Control In tabFilter.Controls ' ENSURE AT LEAST ONE MFR IS SELECTED
            If TypeOf c Is CheckBox AndAlso c.Name.StartsWith("chkMF") Then
                If DirectCast(c, CheckBox).Checked = True Then chksValid = True
            End If
        Next
        If chksValid Then lblChk.Visible = False Else lblChk.Visible = True
        If isNullorNumeric(txtEPmin) AndAlso isNullorNumeric(txtEPmax) Then txtValid = True ' CHECK TEXTBOX VALUES
        If chksValid AndAlso txtValid Then Return True Else Return False ' FINAL CHECK
    End Function

    Private Sub OptionsFound()
        Dim count As Integer = SQL.RecordCount
        If count > 0 Then UpdateObj(lblRecords, count, Color.LawnGreen, Color.DarkSlateGray) Else UpdateObj(lblRecords, count, Color.Red, Color.Gainsboro)
    End Sub

#Region "QUERY FILTERS"
    Private Enum Filter
        mfr
        rpm
        fuel
        burn_type
        nox
        min
        max
        voltage
    End Enum
    Private Function getFilters(Filter As Filter) As String
        Dim FilterList As New List(Of String)
        Dim prefix As String = ""
        Dim partialQuery As String = ""
        Select Case Filter
            Case Filter.mfr : prefix = "chkMF"
            Case Filter.rpm : prefix = "chkER"
            Case Filter.fuel : prefix = "chkFT"
            Case Filter.burn_type : prefix = "chkBT"
            Case Filter.nox : prefix = "chkNX"
            Case Filter.min : prefix = "txtEPmi"
            Case Filter.max : prefix = "txtEPma"
            Case Filter.voltage : prefix = "chkV"
        End Select
        ' ADD FILTERS TO LIST ACCORDING TO TYPE OF FILTER PARAMETER OF FUNCTION
        For Each c As Control In tabFilter.Controls
            If TypeOf c Is CheckBox Then If c.Name.StartsWith(prefix) AndAlso DirectCast(c, CheckBox).Checked = True Then FilterList.Add(c.Text.ToString)
            If TypeOf c Is RadioButton Then If c.Name.StartsWith(prefix) AndAlso DirectCast(c, RadioButton).Checked = True Then FilterList.Add(c.Text.ToString)
            If TypeOf c Is TextBox Then If c.Name.ToString.StartsWith(prefix) AndAlso DirectCast(c, TextBox).Text <> "" Then FilterList.Add(DirectCast(c, TextBox).Text.ToString)
        Next
        ' CHECK TO SEE IF LIST IS EMPTY, IF SO EXIT FUNCTION, ELSE, CONTINUE
        If FilterList.Count < 1 Then Return partialQuery Else partialQuery &= "("
        Select Case Filter ' TODO: MAKE THIS MORE DYNAMIC...
            Case Filter.mfr
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        If f = 0 Then partialQuery &= Filter.ToString & "='" & FilterList(f) & ""
                        If f > 0 Then partialQuery &= "' or " & Filter.ToString & "='" & FilterList(f)
                    Next
                    partialQuery &= "')"
                End If
            Case Filter.rpm
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        If f = 0 Then partialQuery &= Filter.ToString & "=" & FilterList(f)
                        If f > 0 Then partialQuery &= " or " & Filter.ToString & "=" & FilterList(f)
                    Next
                    partialQuery &= ")"
                End If
            Case Filter.fuel : partialQuery &= Filter.ToString & " is null"
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "='" & FilterList(f) & "'"
                        If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "='" & FilterList(f) & "')"
                    Next
                End If
            Case Filter.burn_type : partialQuery &= Filter.ToString & " is null"
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "='" & FilterList(f) & "'"
                        If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "='" & FilterList(f) & "')"
                    Next
                End If
            Case Filter.nox : partialQuery &= Filter.ToString & " is null "
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        '    If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "<=" & FilterList(f)
                        '    If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "<=" & FilterList(f) & ")"
                        Select Case FilterList(f).ToString
                            Case "0.5"
                                If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN 0 AND .8"
                                If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN 0 AND .8)"
                            Case "1.0"
                                If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN .6 AND 1.8"
                                If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN .6 AND 1.8)"
                            Case "2.0"
                                If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN 1.7 AND 3"
                                If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & " BETWEEN 1.7 AND 3)"
                        End Select
                    Next
                End If
            Case Filter.min : partialQuery &= "ElePow100>=" & FilterList(0) & ")"
            Case Filter.max : partialQuery &= "ElePow100<=" & FilterList(0) & ")"
            Case Filter.voltage : partialQuery &= Filter.ToString & " is null"
                If FilterList.Count > 0 Then
                    For f = 0 To FilterList.Count - 1
                        If f < (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "=" & FilterList(f) & ""
                        If f = (FilterList.Count - 1) Then partialQuery &= " or " & Filter.ToString & "=" & FilterList(f) & ")"
                    Next
                End If
        End Select
        FilterList.Clear() ' REFRESH THE LIST
        Return partialQuery
    End Function
#End Region
#End Region

#Region "GENSET CREATION"
    Public Function canConstruct() As Boolean
        Dim RangeErr As String = "percentile must be within range of 30-50..." : Dim NullErr As String = "percentile cannot be left blank..."
        If txtMinExTemp.Text Is Nothing Then txtMinExTemp.Text = 0 : If txtSteam.Text Is Nothing Then txtSteam.Text = 0 : If txtFeed.Text Is Nothing Then txtFeed.Text = 0
        If txtPrimaryInlet.Text Is Nothing Then txtPrimaryInlet.Text = 0 : If txtPrimaryOutlet.Text Is Nothing Then txtPrimaryInlet.Text = 0
        If txt2ndInlet.Text Is Nothing Then txt2ndInlet.Text = 0 : If txt2ndOutlet.Text Is Nothing Then txt2ndOutlet.Text = 0

        ' IF PERCENTILES ARE NULL THEN KICKBACK TO TAB 2 AND ALERT VALUES ARE NULL
        If cbxEngCoolant.SelectedIndex > 0 AndAlso String.IsNullOrEmpty(txtEngCool.Text.ToString) Then MsgBox(String.Format("Engine Coolant {0}", NullErr)) : lblRange.ForeColor = Color.Red : Return False Else 
        If cbxPrimaryCir.SelectedIndex > 0 AndAlso String.IsNullOrEmpty(txtPrimaryCir.Text.ToString) Then MsgBox(String.Format("Primary Circuit {0}", NullErr)) : lblRange.ForeColor = Color.Red : Return False
        If cbx2ndCir.SelectedIndex > 0 AndAlso String.IsNullOrEmpty(txt2ndCir.Text.ToString) Then MsgBox(String.Format("Secondary Circuit {0}", NullErr)) : lblRange.ForeColor = Color.Red : Return False

        ' IF PERCENTILES ARE OUT OF RANGE THEN KICKBACK TO TAB 2 AND ALERT OUT OF BOUNDS VALUES
        If cbxEngCoolant.SelectedIndex > 0 AndAlso Not withinRange(txtEngCool) Then MsgBox(String.Format("Engine Coolant {0}", RangeErr)) : lblRange.ForeColor = Color.Red : Return False
        If cbxPrimaryCir.SelectedIndex > 0 AndAlso Not withinRange(txtPrimaryCir) Then MsgBox(String.Format("Primary Circuit {0}", RangeErr)) : lblRange.ForeColor = Color.Red : Return False
        If cbx2ndCir.SelectedIndex > 0 AndAlso Not withinRange(txt2ndCir) Then MsgBox(String.Format("Secondary Circuit {0}", RangeErr)) : lblRange.ForeColor = Color.Red : Return False
        Return True
    End Function


    Public F1type As Integer : Public F1pct As Double
    Public F2type As Integer : Public F2pct As Double
    Public F3type As Integer : Public F3pct As Double
    Public Sub SetFluids()
        F1type = cbxEngCoolant.SelectedIndex : F2type = cbxPrimaryCir.SelectedIndex
        If cbxEngCoolant.SelectedIndex = 0 Then F1pct = Nothing Else F1pct = txtEngCool.Text
        If cbxPrimaryCir.SelectedIndex = 0 Then F2pct = Nothing Else F2pct = txtPrimaryCir.Text

        If cbx2ndCir.SelectedIndex = -1 Then
            F3type = 3 : F3pct = Nothing
        Else
            F3type = cbx2ndCir.SelectedIndex
            If cbx2ndCir.SelectedIndex > 0 Then F3pct = txt2ndCir.Text
        End If
    End Sub

    Private Function GetLoopCount() As Integer
        If radSelected.Checked Then : Return 1
        ElseIf radTop5.Checked Then : Return 5
        ElseIf radTop10.Checked Then : Return 10
        ElseIf radAll.Checked Then : Return SQL.RecordCount : End If
        Return Nothing
    End Function

    Private Sub ConstructGenset()
        Me.Cursor = Cursors.WaitCursor
        If canConstruct() Then
            Dim TimerStart As DateTime = Now
            Dim TimeSpent As System.TimeSpan
            Dim loopCount As Integer = GetLoopCount()
            MsgBox(String.Format("Loop Count = {0}", loopCount))
            SetFluids()

            If loopCount = 1 Then
                Index = dgvCompare.CurrentRow.Index
                Try
                    MyGenset = New Genset(_get(SQL.DBDS, "id", Index), _get(SQL.DBDS, "mfr", Index), _get(SQL.DBDS, "model", Index), _get(SQL.DBDS, "rpm", Index), _get(SQL.DBDS, "fuel", Index), _
                                        _get(SQL.DBDS, "elepow100"), PowFactor, _
                                        CDbl(txtMinExTemp.Text), CDbl(txtSteam.Text), CDbl(txtFeed.Text), CDbl(txtPrimaryInlet.Text), CDbl(txtPrimaryOutlet.Text), _
                                        CDbl(txt2ndInlet.Text), CDbl(txt2ndOutlet.Text), chkSteam.Checked, chkEhru.Checked, radEHRUtoJW.Checked, radEHRUtoPrimary.Checked, _
                                        chkRecoverJW.Checked, chkRecoverLT.Checked, radAddToPrimary.Checked, radAddTo2nd.Checked, F1type, F2type, F3type, F1pct, F2pct, F3pct, _
                                         radOilToJw.Checked, radOilToIc.Checked)
                Catch ex As Exception
                    MsgBox(ex.Message & Environment.NewLine & ex.ToString)
                End Try
            Else
                ' USE A WHILE LOOP HERE
                'Index = 0
                'For Each r As DataRow In SQL.DBDS.Tables(Index).Rows
                Try
                    MyGenset = New Genset(_get(SQL.DBDS, "id"), _get(SQL.DBDS, "mfr"), _get(SQL.DBDS, "model"), _get(SQL.DBDS, "rpm"), _get(SQL.DBDS, "fuel"), _get(SQL.DBDS, "elepow100"), PowFactor, _
                                      CDbl(txtMinExTemp.Text), CDbl(txtSteam.Text), CDbl(txtFeed.Text), CDbl(txtPrimaryInlet.Text), CDbl(txtPrimaryOutlet.Text), _
                                      CDbl(txt2ndInlet.Text), CDbl(txt2ndOutlet.Text), chkSteam.Checked, chkEhru.Checked, radEHRUtoJW.Checked, radEHRUtoPrimary.Checked, _
                                      chkRecoverJW.Checked, chkRecoverLT.Checked, radAddToPrimary.Checked, radAddTo2nd.Checked, F1type, F2type, F3type, F1pct, F2pct, F3pct, _
                                      radOilToJw.Checked, radOilToIc.Checked)
                Catch ex As Exception
                    MsgBox(ex.Message & Environment.NewLine & ex.ToString)
                End Try
                'Next

            End If

            TimeSpent = Now.Subtract(TimerStart) : MsgBox(String.Format("Time spent = {0:n3} seconds", TimeSpent.TotalSeconds))
        Else
            MsgBox("Something went wrong")
        End If
        Me.Cursor = Cursors.Default
    End Sub
#End Region

#Region "FORM OBJECTS"
#Region "TabControl \ Filter"
    Private Sub DynamicSynthQuery(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMFgua.CheckedChanged, chkMFmtu.CheckedChanged, chkER1200.CheckedChanged, chkER1500.CheckedChanged, chkER1800.CheckedChanged, _
        chkFT1.CheckedChanged, chkFT2.CheckedChanged, chkFT3.CheckedChanged, chkFT4.CheckedChanged, chkFT5.CheckedChanged, chkFT6.CheckedChanged, chkBTrich.CheckedChanged, chkNX1.CheckedChanged, chkNX2.CheckedChanged, chkNX3.CheckedChanged, _
        txtEPmin.TextChanged, txtEPmax.TextChanged, chkV480.CheckedChanged, chkV600.CheckedChanged, chkV4160.CheckedChanged, chkV12470.CheckedChanged, radPF8.CheckedChanged, radPF9.CheckedChanged, radPF1.CheckedChanged
        ' IF ANY OF THE ABOVE OBJECTS CHANGED DURING RUNTIME, PERFORM SynthQuery()
        If CType(sender, Control).Name.StartsWith("rad") Then
            Select Case DirectCast(sender, RadioButton).Name.ToString
                Case "radPF8" : PowFactor = 0.8
                Case "radPF9" : PowFactor = 0.9
                Case "radPF1" : PowFactor = 1
            End Select
        End If
        SynthQuery()
    End Sub

    ' OBJECT BEHAVIOR
    Private Sub CheckBox_ToggleEnabled(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMfr.CheckedChanged, chkRpm.CheckedChanged, chkFuel.CheckedChanged, chkEmit.CheckedChanged, chkOlts.CheckedChanged
        If DirectCast(sender, CheckBox).Checked = True Then ToggleEnabled((DirectCast(sender, CheckBox).Name.ToString), False, tabFilter) Else ToggleEnabled((DirectCast(sender, CheckBox).Name.ToString), True, tabFilter)
    End Sub
    Private Sub FilterTab_TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEPmin.TextChanged, txtEPmax.TextChanged
        If Not isNullorNumeric(sender) Then DirectCast(sender, TextBox).ResetText()
    End Sub
    Private Sub chkBTlean_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBTlean.CheckedChanged
        If chkBTlean.Checked = True Then : lblNox.Visible = True : chkNX1.Visible = True : chkNX2.Visible = True : chkNX3.Visible = True : SynthQuery()
        Else
            lblNox.Visible = False : chkNX1.Visible = False : chkNX2.Visible = False : chkNX3.Visible = False : chkNX1.Checked = False : chkNX2.Checked = False : chkNX3.Checked = False : SynthQuery()
        End If
    End Sub
    Private Sub ToggleEnabled(ByVal Panel As String, ByVal Toggle As Boolean, ByVal Parent As Object)
        Dim strPrefix As String = ""
        Select Case Panel
            Case "chkMfr" : strPrefix = "chkMF" ' engineMFR
            Case "chkRpm" : strPrefix = "chkER" ' engineRPM
            Case "chkFuel" : strPrefix = "chkFT" ' fuelType
            Case "chkEmit" : strPrefix = "chkBT" ' burnType (emissions)
            Case "chkOlts" : strPrefix = "chkV" ' voltage
            Case "chkPowAny" : strPrefix = "chkPF" ' power factor
        End Select
        ' BEGIN TOGGLE
        For Each c As Control In Parent.controls
            If c.Name.StartsWith(strPrefix) Then DirectCast(c, CheckBox).Enabled = Toggle
            If Toggle Then
                If c.Name.StartsWith(strPrefix) Then DirectCast(c, CheckBox).Checked = False
            Else
                If c.Name.StartsWith(strPrefix) Then DirectCast(c, CheckBox).Checked = True
            End If
        Next
    End Sub
#End Region
#Region "TabControl \ Heat Recovery"
    Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinExTemp.TextChanged, txtSteam.TextChanged, txtFeed.TextChanged, txtPrimaryInlet.TextChanged, txtPrimaryOutlet.TextChanged, txt2ndInlet.TextChanged, txt2ndOutlet.TextChanged
        If Not isNotNullorInvalid(sender) Then DirectCast(sender, TextBox).Text = Nothing
    End Sub
    ' FLUID PERCENTILES
    Private Sub CircuitFluid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEngCool.TextChanged, txtPrimaryCir.TextChanged, txt2ndCir.TextChanged
        If isNotNullorInvalid(sender) Then
            If withinRange(sender) Then lblRange.ForeColor = Color.Black Else lblRange.ForeColor = Color.Red
        Else : DirectCast(sender, TextBox).ResetText() : lblRange.ForeColor = Color.Black : End If
    End Sub
    ' CHECKBOXES (LEFT HALF)
    Private Sub chkSteam_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSteam.CheckedChanged
        If chkSteam.Checked = True Then ToggleHeatControls(txtSteam, txtFeed, True, False, True) Else ToggleHeatControls(txtSteam, txtFeed, False, False, True) : txtSteam.Text = 0 : txtFeed.Text = 0
    End Sub
    Private Sub chkEhru_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEhru.CheckedChanged
        If chkEhru.Checked Then : radEHRUtoJW.Checked = True : ToggleHeatControls(radEHRUtoJW, radEHRUtoPrimary, True, False, True) : Else : ToggleHeatControls(radEHRUtoJW, radEHRUtoPrimary, False, True, True) : End If
    End Sub
    ' CHECKBOXES (RIGHT HALF)
    Private Sub chkRecoverJW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecoverJW.CheckedChanged
        If chkRecoverJW.Checked Then txtPrimaryInlet.Focus() : txtPrimaryInlet.SelectAll()
    End Sub
    Private Sub chkRecoverLT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecoverLT.CheckedChanged
        If chkRecoverLT.Checked Then : radAddTo2nd.Checked = True : ToggleHeatControls(radAddToPrimary, radAddTo2nd, True, False, True) : ToggleHeatControls(txt2ndInlet, txt2ndOutlet, True, False, True) : txt2ndInlet.Focus() : txt2ndOutlet.SelectAll()
        Else : ToggleHeatControls(radAddToPrimary, radAddTo2nd, False, True, True) : ToggleHeatControls(txt2ndInlet, txt2ndOutlet, False, False, True, True) : End If
    End Sub
    Private Sub radAddTo2nd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radAddTo2nd.CheckedChanged
        If radAddTo2nd.Checked Then : cbx2ndCir.Enabled = True : cbx2ndCir.SelectedIndex = 0 : txt2ndInlet.SelectAll()
        Else : cbx2ndCir.Enabled = False : cbx2ndCir.SelectedIndex = -1 : txt2ndInlet.Text = 0 : txt2ndOutlet.Text = 0 : End If
    End Sub
    ' COMBOBOXES
    Private Sub cbxEngCoolant_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxEngCoolant.SelectedIndexChanged
        If cbxEngCoolant.SelectedIndex > 0 Then txtEngCool.Visible = True Else txtEngCool.ResetText() : txtEngCool.Visible = False
    End Sub
    Private Sub cbxPrimaryCir_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxPrimaryCir.SelectedIndexChanged
        If cbxPrimaryCir.SelectedIndex > 0 Then txtPrimaryCir.Visible = True Else txtPrimaryCir.ResetText() : txtPrimaryCir.Visible = False
    End Sub
    Private Sub cbx2ndCir_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbx2ndCir.SelectedIndexChanged
        If cbx2ndCir.SelectedIndex > 0 Then txt2ndCir.Visible = True Else txt2ndCir.ResetText() : txt2ndCir.Visible = False
    End Sub
#End Region
#Region "TabControl \ Compare"
    Private Sub EngineMode()
        dgvGensets.Visible = False : pnlEngines.Visible = True : dgvCompare.Visible = True : UpdateObj(lblMode, "ENGINE MODE", Color.Black) : pnlMode.BackColor = Color.Transparent
    End Sub
    Private Sub GensetMode()
        dgvCompare.Visible = False : pnlEngines.Visible = False : dgvGensets.Visible = True : UpdateObj(lblMode, "GENSET MODE", Color.Chartreuse) : pnlMode.BackColor = Color.DarkSlateGray
    End Sub
    Private Sub radEngines_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles radEngines.CheckedChanged
        If radEngines.Checked Then EngineMode() Else GensetMode()
    End Sub
    Private Sub radGensets_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles radGensets.CheckedChanged
        If radGensets.Checked = True Then GensetMode() Else EngineMode()
    End Sub

    ' WILDCARD SEARCH
    Private Sub txtSearch_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSearch.TextChanged, cbxFilter.SelectedIndexChanged
        If txtSearch.Text <> "" Then Search(txtSearch.Text, cbxFilter.Text) Else SynthQuery()
    End Sub
    Private Sub Search(key As String, filter As String)
        'If SQL.RecordCount > 0 Then SQL.DBDS.Clear()
        SQL.AddParam("@key", String.Format("%{0}%", key))
        SQL.ExecQuery(String.Format("SELECT id, mfr, model, rpm, fuel, burn_type, nox, elepow100, voltage FROM Engines WHERE {0} LIKE @key ORDER BY elepow100, rpm", filter))
        If Not String.IsNullOrEmpty(SQL.Exception) Then MsgBox(SQL.Exception) : Exit Sub
        dgvCompare.DataSource = SQL.DBDS.Tables(0) : lblTotal.Text = SQL.RecordCount
        'If GensetList.Count > 0 Then GensetList.Clear()
    End Sub
    Private Sub btnWipe_Click(sender As System.Object, e As System.EventArgs) Handles btnWipe.Click
        txtSearch.ResetText()
    End Sub

    ' BEGIN GENSET CREATION
    Private Sub btnPopulate_Click(sender As System.Object, e As System.EventArgs) Handles btnPopulate.Click
        If SQL.RecordCount > 0 Then ConstructGenset()
        'PopulateList()
        'radGensets.Enabled = True : radGensets.Checked = True
    End Sub

    

    Private Sub PopulateList()
        Me.Cursor = Cursors.WaitCursor
        ' MANAGE PROGRESS BAR AND TIMER
        prgMain.Value = 0 : prgMain.Maximum = SQL.RecordCount : prgMain.Visible = True
        Dim TimerStart As DateTime = Now : Dim TimeSpent As System.TimeSpan

        ' CLEAR CURRENT ROWS/LIST AND ALLOCATE NEW ROWS
        dgvGensets.Rows.Clear() : GensetList.Clear()
        dgvGensets.Rows.Add(SQL.RecordCount)

        ' SET FLUID PERCENTAGE VALUES
        Dim F1type As Integer = cbxEngCoolant.SelectedIndex : Dim F1pct As Integer : If cbxEngCoolant.SelectedIndex > 0 Then F1pct = txtEngCool.Text
        Dim F2type As Integer = cbxPrimaryCir.SelectedIndex : Dim F2pct As Integer : If cbxPrimaryCir.SelectedIndex > 0 Then F2pct = txtPrimaryCir.Text
        Dim F3type As Integer : Dim F3pct As Integer
        If cbx2ndCir.SelectedIndex = -1 Then
            F3type = 3
        Else
            F3type = cbx2ndCir.SelectedIndex
            If cbx2ndCir.SelectedIndex > 0 Then F3pct = txt2ndCir.Text
        End If

        ' BEGIN GENSET LIST CREATION
        Dim i As Integer = 0
        For Each r As DataRow In SQL.DBDS.Tables(i).Rows
            prgMain.Value += 1
            MyGenset = New Genset(_get(SQL.DBDS, "id"), _get(SQL.DBDS, "mfr"), _get(SQL.DBDS, "model"), _get(SQL.DBDS, "rpm"), _get(SQL.DBDS, "fuel"), _get(SQL.DBDS, "elepow100"), PowFactor, _
                                  CDbl(txtMinExTemp.Text), CDbl(txtSteam.Text), CDbl(txtFeed.Text), CDbl(txtPrimaryInlet.Text), CDbl(txtPrimaryOutlet.Text), _
                                  CDbl(txt2ndInlet.Text), CDbl(txt2ndOutlet.Text), chkSteam.Checked, chkEhru.Checked, radEHRUtoJW.Checked, radEHRUtoPrimary.Checked, _
                                  chkRecoverJW.Checked, chkRecoverLT.Checked, radAddToPrimary.Checked, radAddTo2nd.Checked, F1type, F2type, F3type, F1pct, F2pct, F3pct, _
                                  radOilToJw.Checked, radOilToIc.Checked)
            ' OUTPUT INFO
            'lblCase.Text = MyGenSet.CalcCase : PrintAllStats()

            'GensetList.Add(MyGenset)

            ' POPULATE ROWS
            'With MyGenset
            '    dgvGensets.Rows(i).Cells(0).Value = ._EngineID : dgvGensets.Rows(i).Cells(1).Value = ._MFR : dgvGensets.Rows(i).Cells(2).Value = ._Model : dgvGensets.Rows(i).Cells(3).Value = .RPM
            '    dgvGensets.Rows(i).Cells(4).Value = ._Fuel : dgvGensets.Rows(i).Cells(5).Value = .EngKW : dgvGensets.Rows(i).Cells(6).Value = ._LTheat100u : dgvGensets.Rows(i).Cells(7).Value = ._FuelCon100u
            '    dgvGensets.Rows(i).Cells(8).Value = .bHPhr : dgvGensets.Rows(i).Cells(9).Value = .QSteam : dgvGensets.Rows(i).Cells(10).Value = ._HeatMain100u : dgvGensets.Rows(i).Cells(11).Value = .QEHRU
            '    dgvGensets.Rows(i).Cells(12).Value = ._OilCool100u : dgvGensets.Rows(i).Cells(13).Value = .QHX : dgvGensets.Rows(i).Cells(14).Value = .QICHX : dgvGensets.Rows(i).Cells(15).Value = String.Format("{0:n1}", .EleEff)
            '    dgvGensets.Rows(i).Cells(16).Value = String.Format("{0:n1}", .ThermEff) : dgvGensets.Rows(i).Cells(17).Value = String.Format("{0:n1}", .TotalEff) : dgvGensets.Rows(i).Cells(18).Value = .PwFlow : dgvGensets.Rows(i).Cells(19).Value = .PwInActual
            '    dgvGensets.Rows(i).Cells(20).Value = .PwOutActual : dgvGensets.Rows(i).Cells(21).Value = .SWFlow : dgvGensets.Rows(i).Cells(22).Value = .SwInActual : dgvGensets.Rows(i).Cells(23).Value = .SwOutActual
            'End With
            i += 1
        Next

        ' FINALIZE PROGRESS BAR AND REPORT TIME SPENT
        TimeSpent = Now.Subtract(TimerStart) : MsgBox(String.Format("Time spent = {0:n3} seconds", TimeSpent.TotalSeconds)) : prgMain.Visible = False
        Me.Cursor = Cursors.Default
    End Sub
#End Region
#Region "TabControl \ View"

#End Region
#Region "TabControl \ Final"

#End Region
#Region "Buttons"
    Private Sub NavButtons_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click, btnBack.Click
        Navigate(DirectCast(sender, Button).Name, tcMain, lblRecords.Text)
    End Sub
#End Region
#Region "Menu Strip"
    Private Sub miDB_Click(sender As System.Object, e As System.EventArgs) Handles miDB.Click
        Dim EditDB As New frmEditor : EditDB.Show()
    End Sub
#End Region
#End Region
End Class