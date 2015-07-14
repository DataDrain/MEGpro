Imports System.Drawing
Imports System.Threading

Public Class frmMain
    Public Version As String = "1.5.0"
#Region "DECLARATIONS"
    Private SQL As New SQLControl
    Public PowFactor As Single = 1
    Public eng_index As Integer : Public genset_index As Integer

    Private MyGenset As Genset
    Private GensetList As New List(Of Genset)

    Public bmp As Bitmap
#End Region

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        InitializeForm()
    End Sub

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
        lblTotal.Text = SQL.RecordCount
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
        'If cbx2ndCir.SelectedIndex = -1 Then : F3type = 3 : F3pct = Nothing : Else : F3type = cbx2ndCir.SelectedIndex : If cbx2ndCir.SelectedIndex > 0 Then F3pct = txt2ndCir.Text : End If
    End Sub

    Private Function GetLoopCount() As Integer
        If radSelected.Checked Then : Return 1
        ElseIf radTop5.Checked Then : If SQL.RecordCount < 5 Then Return SQL.RecordCount Else Return 5
        ElseIf radTop10.Checked Then : If SQL.RecordCount < 10 Then Return SQL.RecordCount Else Return 10
        ElseIf radTop20.Checked Then : If SQL.RecordCount < 20 Then Return SQL.RecordCount Else Return 20
        ElseIf radAll.Checked Then : Return SQL.RecordCount : End If
        Return Nothing
    End Function

    Private Sub ConsructionProcess()
        Me.Cursor = Cursors.WaitCursor
        Dim TimerStart As DateTime = Now : Dim TimeSpent As System.TimeSpan
        Dim loopCount As Integer = GetLoopCount()
        dgvGensets.Rows.Clear() : GensetList.Clear() : dgvGensets.Rows.Add(loopCount)

        If canConstruct() Then
            SetFluids()
            If loopCount = 1 Then
                Try
                    ConstructGenset(eng_index)
                Catch ex As Exception
                    MsgBox(ex.Message & Environment.NewLine & ex.ToString)
                End Try
                tcMain.SelectedIndex += 1 : lblCurrent.Text = eng_index + 1 : lblTotal.Text = SQL.RecordCount
            Else
                eng_index = 0
                While eng_index < loopCount
                    Try
                        ConstructGenset(eng_index)
                    Catch ex As Exception
                        MsgBox(ex.Message & Environment.NewLine & ex.ToString)
                    End Try
                    GensetList.Add(MyGenset)
                    PopulateGensetDGV(eng_index)
                    eng_index += 1
                End While
                radGensets.Checked = True : lblCurrent.Text = dgvGensets.CurrentRow.Index + 1 : lblTotal.Text = loopCount
            End If
        Else
            MsgBox("Something went wrong")
        End If
        TimeSpent = Now.Subtract(TimerStart) : lblTimer.Text = String.Format("{0:n3} sec", TimeSpent.TotalSeconds) : Me.Cursor = Cursors.Default
        PrintAllStats()
    End Sub

    Private Sub ConstructGenset(i As Integer)
        MyGenset = New Genset(dgvCompare.Rows(i).Cells("id").Value, dgvCompare.Rows(i).Cells("mfr").Value, _
                            CDbl(txtMinExTemp.Text), CDbl(txtSteam.Text), CDbl(txtFeed.Text), CDbl(txtPrimaryInlet.Text), CDbl(txtPrimaryOutlet.Text), _
                            CDbl(txt2ndInlet.Text), CDbl(txt2ndOutlet.Text), chkSteam.Checked, chkEhru.Checked, radEHRUtoJW.Checked, radEHRUtoPrimary.Checked, _
                            chkRecoverJW.Checked, chkRecoverLT.Checked, radAddToPrimary.Checked, radAddTo2nd.Checked, F1type, F2type, F3type, F1pct, F2pct, F3pct, _
                            radOilToJw.Checked, radOilToIc.Checked)
    End Sub
    Private Sub PopulateGensetDGV(i As Integer)
        With MyGenset
            dgvGensets.Rows(i).Cells(0).Value = ._EngID : dgvGensets.Rows(i).Cells(1).Value = ._MFR : dgvGensets.Rows(i).Cells(2).Value = ._Model : dgvGensets.Rows(i).Cells(3).Value = ._RPM
            dgvGensets.Rows(i).Cells(4).Value = ._Fuel : dgvGensets.Rows(i).Cells(5).Value = .KWeOut100 : dgvGensets.Rows(i).Cells(6).Value = .lt_heat100 : dgvGensets.Rows(i).Cells(7).Value = .fuelcon100
            dgvGensets.Rows(i).Cells(8).Value = .bHPhr : dgvGensets.Rows(i).Cells(9).Value = .QSteam : dgvGensets.Rows(i).Cells(10).Value = .mainheat100 : dgvGensets.Rows(i).Cells(11).Value = .QEHRU
            dgvGensets.Rows(i).Cells(12).Value = .oilcool100 : dgvGensets.Rows(i).Cells(13).Value = .QHX : dgvGensets.Rows(i).Cells(14).Value = .QICHX : dgvGensets.Rows(i).Cells(15).Value = OneDec(.EleEff)
            dgvGensets.Rows(i).Cells(16).Value = OneDec(.ThermEff) : dgvGensets.Rows(i).Cells(17).Value = OneDec(.TotalEff) : dgvGensets.Rows(i).Cells(18).Value = .PwFlow : dgvGensets.Rows(i).Cells(19).Value = .PwInActual
            dgvGensets.Rows(i).Cells(20).Value = .PwOutActual : dgvGensets.Rows(i).Cells(21).Value = .SWFlow : dgvGensets.Rows(i).Cells(22).Value = .SwInActual : dgvGensets.Rows(i).Cells(23).Value = .SwOutActual
        End With
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
        dgvGensets.Visible = False : pnlEngines.Visible = True : dgvCompare.Visible = True : UpdateObj(lblMode, "ENGINE COMPARISON MODE", Color.Black) : pnlMode.BackColor = Color.Transparent
    End Sub
    Private Sub GensetMode()
        dgvCompare.Visible = False : pnlEngines.Visible = False : dgvGensets.Visible = True : UpdateObj(lblMode, "GENSET COMPARISON MODE", Color.Chartreuse) : pnlMode.BackColor = Color.DarkSlateGray
    End Sub
    Private Sub radEngines_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles radEngines.CheckedChanged
        If radEngines.Checked Then EngineMode() Else GensetMode()
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
        If SQL.RecordCount > 0 Then ConsructionProcess()
    End Sub

    ' HANDLE INDEX CHANGES
    Private Sub dgvCompare_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvCompare.SelectionChanged
        GetMyEngine()
    End Sub
    Private Sub dgvCompare_Sorted(sender As Object, e As System.EventArgs) Handles dgvCompare.Sorted
        dgvCompare.ClearSelection() : dgvCompare.CurrentCell = Nothing : GetMyEngine()
    End Sub
    Private Sub GetMyEngine()
        Try
            eng_index = dgvCompare.CurrentRow.Index
        Catch ex As Exception : End Try ' DO NOTHING
    End Sub

    ' HANDLE GENSET DGV
    Private Sub dgvGensets_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvGensets.SelectionChanged
        GetMyGenset()
        'MyGenset = GensetList.Find()
    End Sub
    Private Sub dgvGensets_Sorted(sender As Object, e As System.EventArgs) Handles dgvGensets.Sorted
        GetMyGenset()
    End Sub
    Private Sub GetMyGenset()
        Dim gensetID As String = Nothing
        Try
            genset_index = dgvGensets.CurrentRow.Index
            gensetID = dgvGensets.Rows(genset_index).Cells(0).Value
            ' SEARCH LIST OF GENSETS
            MyGenset = GensetList.Find(Function(x) x._EngID = gensetID)
            PrintAllStats()
        Catch ex As Exception : End Try ' DO NOTHING
    End Sub
#End Region
#Region "TabControl \ View"
    Private Sub btnRight_Click(sender As System.Object, e As System.EventArgs) Handles btnRight.Click
        NaviageGensetSpecs("next")
    End Sub
    Private Sub btnLeft_Click(sender As System.Object, e As System.EventArgs) Handles btnLeft.Click
        NaviageGensetSpecs("prev")
    End Sub
    Private Sub NaviageGensetSpecs(direction As String)
        Dim listPos As Integer = GensetList.IndexOf(MyGenSet)
        If GensetList.Count > 0 Then
            Select Case direction
                Case "prev"
                    If listPos = 0 Then listPos = GensetList.Count - 1 Else listPos -= 1
                Case "next"
                    If listPos = GensetList.Count - 1 Then listPos = 0 Else listPos += 1
            End Select
            MyGenSet = GensetList(listPos)
            PrintAllStats()
            lblCurrent.Text = listPos + 1
        Else
            MsgBox("There is not a list of Gensets created yet...")
        End If
    End Sub
#End Region
#Region "TabControl \ Final"

#End Region

#Region "Menu Strip"
    Private Sub miDB_Click(sender As System.Object, e As System.EventArgs) Handles miDB.Click
        Dim EditDB As New frmEditor : EditDB.Show()
    End Sub
#End Region
#End Region

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

    Private Sub NavButtons_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click, btnBack.Click
        Navigate(DirectCast(sender, Button).Name, tcMain, lblRecords.Text)
    End Sub

    Public Sub PrintAllStats()
        Try
            tabView.AutoScrollPosition = New Point(0, 0)
            'If MyGenSet.tooHot = True Then : lblWarning.Visible = True : Else : lblWarning.Visible = False : End If
            With MyGenset
                ' SET VIEW MODE (PARTIAL POWER LOADS)
                If ._MFR = "Guascor" Then ViewMode("20", tabView) Else ViewMode("25", tabView)
                Dim myToolTipTxt = NoDec(.QHX) : ToolTip.SetToolTip(lblPrimHeat, myToolTipTxt) : lblCase.Text = .CalcCase
                '============ TOP PANEL (GENERAL INFO) ========================
                lblMFR.Text = String.Format("{0} - {1}", ._MFR, ._Model) : lblEngineID.Text = ._EngID : lblFuel.Text = ._Fuel : lblGenID.Text = ._genID
                lblRPM.Text = ._RPM : lblKW.Text = String.Format("{0:n0} KW", .KWeOut100) : lblVolts.Text = String.Format("{0} Volts", ._genVolts) : lblPowFactor.Text = PowFactor
                '============ ENGINE PERFORMANCE PANEL ========================
                lblKWE100.Text = NoDec(.KWeOut100)
                lblBHP100.Text = .engpow100
                lblExFlow100.Text = NoDec(.exflow100)
                lblExTemp100.Text = NoDec(.extemp100)
                lblHeatMain100.Text = NoDec(.mainheat100)
                lblQExAvail.Text = NoDec(.QExAvail)
                lblLTheat100.Text = NoDec(.lt_heat100)
                lblJWin.Text = NoDec(.jw_in) : lblJWout.Text = .jw_out : lblJWFlowRate.Text = NoDec(.jw_flow)
                lblICin.Text = .ic_in : lblICout.Text = NoDec(.ic_out) : lblICFlowRate.Text = NoDec(.ic_flow)
                lblOilCool100.Text = NoDec(.oilcool100)
                lblFuelConHr100.Text = NoDec(.fuelcon100)
                lblFuelConBhp100.Text = NoDec(.bHPhr)
                lblFuelKW100.Text = NoDec(.btuKWh)
                '============ RECOVERED HEAT PANEL ========================
                lblJWtoPrimary100.Text = NoDec(.mainheat100)
                lblExRecov100.Text = NoDec(.QEHRU)
                If ._MFR = "Guascor" Then lblOiltoPrimary.Text = NoDec(.oilcool100) Else lblOiltoPrimary.Text = 0
                .QJWRad *= -1 : lblJWRad100.Text = NoDec(.QJWRad)
                lblPrimHeat.Text = NoDec((CInt(lblJWtoPrimary100.Text) + CInt(lblExRecov100.Text) + CInt(lblOiltoPrimary.Text) + CInt(lblJWRad100.Text)))
                If Math.Abs(CDbl(lblPrimHeat.Text) - .QHX) > 5000 Then lblStatus.Text = "Total Primary Heat was adjusted, QHX greater than +/- 5,000." Else lblStatus.Text = ""
                lblQICHX100.Text = NoDec(.QICHX)
                .QICRad *= -1 : lblICRad100.Text = NoDec(.QICRad)
                lblSecHeat100.Text = NoDec(CDbl(lblQICHX100.Text) + CDbl(lblICRad100.Text))
                lblEleEff100.Text = OneDec(.EleEff)
                lblThermEff100.Text = OneDec(.ThermEff)
                lblTotalEff100.Text = OneDec(.TotalEff)
                '============ PRIMARY CIRCUIT PANEL ========================
                lblPWFlow100.Text = NoDec(.PwFlow)
                lblPwInAct100.Text = NoDec(.PwInActual)
                lblPwOutAct100.Text = NoDec(.PwOutActual)
                lblFluid2Type.Text = ._PrmCir_fluid.ToString : lblFluid2Percent.Text = String.Format("{0}%", ._f2pct)
                '============ SECONDARY CIRCUIT PANEL ========================
                lblSWFlow100.Text = NoDec(.SWFlow)
                lblSwInAct100.Text = NoDec(.SwInActual)
                lblSwOutAct100.Text = NoDec(.SwOutActual)
                lblFluid3Type.Text = ._SecCir_fluid.ToString : lblFluid3Percent.Text = String.Format("{0}%", ._f3pct)
                '============ STEAM PANEL ========================
                lblQsteam100.Text = NoDec(.QSteam)
                lblSteamProd100.Text = NoDec(.SteamProduction)
                lblSteamPress.Text = NoDec(._user_StmPress)
                Select Case ._MFR '============ PARTIAL PANELS ========================
                    Case "MTU"
                        '============ ENGINE PERFORMANCE PANEL ========================
                        lblKWE75.Text = NoDec(.KWeOut75) : lblKWE50.Text = NoDec(.KWeOut50)
                        lblBHP75.Text = .engpow75 : lblBHP50.Text = .engpow50
                        lblExFlow75.Text = NoDec(.exflow75) : lblExFlow50.Text = NoDec(.exflow50)
                        lblExTemp75.Text = NoDec(.extemp75) : lblExTemp50.Text = NoDec(.extemp50)
                        lblHeatMain75.Text = NoDec(.mainheat75) : lblHeatMain50.Text = NoDec(.mainheat50)
                        lblQExAvail75.Text = NoDec(.QExAvail75) : lblQExAvail50.Text = NoDec(.QExAvail50)
                        lblLTheat75.Text = NoDec(.lt_heat75) : lblLTheat50.Text = NoDec(.lt_heat50)
                        lblOilCool75.Text = 0 : lblOilCool50.Text = 0 ' set to 0 because this is guascor only
                        lblJWin75.Text = .jw_in : lblJWin50.Text = .jw_in
                        lblJwOut75.Text = NoDec(.jwout75) : lblJwOut50.Text = NoDec(.jwout50)
                        lblJWFlowRate75.Text = NoDec(.jw_flow) : lblJWFlowRate50.Text = NoDec(.jw_flow)
                        lblICin75.Text = .ic_in : lblICin50.Text = .ic_in
                        lblICout75.Text = NoDec(.icout75) : lblICout50.Text = NoDec(.icout50)
                        lblICFlowRate75.Text = NoDec(.ic_flow) : lblICFlowRate50.Text = NoDec(.ic_flow)
                        lblFuelConHr75.Text = NoDec(.fuelcon75) : lblFuelConHr50.Text = NoDec(.fuelcon50)
                        lblFuelConBhp75.Text = NoDec(.bHPhr75) : lblFuelConBhp50.Text = NoDec(.bHPhr50)
                        lblFuelKW75.Text = NoDec(.btuKWh75) : lblFuelKW50.Text = NoDec(.btuKW50)
                        '============ RECOVERED HEAT PANEL ========================
                        lblJWtoPrimary75.Text = NoDec(.mainheat75) : lblJWtoPrimary50.Text = NoDec(.mainheat50)
                        lblExRecov75.Text = NoDec(.QEHRU75) : lblExRecov50.Text = NoDec(.QEHRU50)
                        .QJWRad75 *= -1 : lblJWRad75.Text = NoDec(.QJWRad75) : .QJWRad50 *= -1 : lblJWRad50.Text = NoDec(.QJWRad50)
                        lblPrimHeat75.Text = NoDec(CDbl(lblJWtoPrimary75.Text) + CDbl(lblExRecov75.Text) + CDbl(lblOilCool75.Text) + CDbl(lblJWRad75.Text))
                        lblPrimHeat50.Text = NoDec(CDbl(lblJWtoPrimary50.Text) + CDbl(lblExRecov50.Text) + CDbl(lblOilCool50.Text) + CDbl(lblJWRad50.Text))
                        lblQICHX75.Text = NoDec(.QICHX75) : lblQICHX50.Text = NoDec(.QICHX50)
                        .QICRad75 *= -1 : lblICRad75.Text = NoDec(.QICRad75) : .QICRad50 *= -1 : lblICRad50.Text = NoDec(.QICRad50)
                        lblSecHeat75.Text = NoDec(CDbl(lblQICHX75.Text) + CDbl(lblICRad75.Text)) : lblSecHeat50.Text = NoDec(CDbl(lblQICHX50.Text) + CDbl(lblICRad50.Text))
                        lblEleEff75.Text = OneDec(.EleEff75) : lblEleEff50.Text = OneDec(.EleEff50)
                        lblThermEff75.Text = OneDec(.ThermEff75) : lblThermEff50.Text = OneDec(.ThermEff50)
                        lblTotalEff75.Text = OneDec(.TotalEff75) : lblTotalEff50.Text = OneDec(.TotalEff50)
                        '============ PRIMARY CIRCUIT PANEL ========================
                        lblPWFlow75.Text = NoDec(.PwFlow75) : lblPWFlow50.Text = NoDec(.PwFlow50)
                        lblPwInAct75.Text = NoDec(.PwInActual75) : lblPwInAct50.Text = NoDec(.PwInActual50)
                        lblPwOutAct75.Text = NoDec(.PwOutActual75) : lblPwOut50.Text = NoDec(.PwOutActual50)
                        lblJWRad75.Text = NoDec(.QJWRad75) : lblJWRad50.Text = NoDec(.QJWRad50)
                        '============ SECONDARY CIRCUIT PANEL ========================
                        lblSWFlow75.Text = NoDec(.SwFlow75) : lblSWFlow50.Text = NoDec(.SwFlow50)
                        lblSwInAct75.Text = NoDec(.SwInActual75) : lblSwInAct50.Text = NoDec(.SwInActual50)
                        lblSwOutAct75.Text = NoDec(.SwOutActual75) : lblSwOutAct50.Text = NoDec(.SwOutActual50)
                        lblICRad75.Text = NoDec(.QICRad75) : lblICRad50.Text = NoDec(.QICRad50)
                        '============ STEAM PANEL ========================
                        lblQsteam75.Text = NoDec(.QSteam75) : lblQsteam50.Text = NoDec(.QSteam50)
                        lblSteamProd75.Text = NoDec(.SteamProd75) : lblSteamProd50.Text = NoDec(.SteamProd50)
                    Case "MAN"
                    Case "Guascor"
                        '============ ENGINE PERFORMANCE PANEL ========================
                        lblKWE80.Text = NoDec(.KWeOut80) : lblKWE60.Text = NoDec(.KWeOut60) : lblKWE40.Text = NoDec(.KWeOut40)
                        lblBHP80.Text = .engpow80 : lblBHP60.Text = .engpow60 : lblBHP40.Text = .engpow40
                        lblExFlow80.Text = NoDec(.exflow80) : lblExFlow60.Text = NoDec(.exflow60) : lblExFlow40.Text = NoDec(.exflow40)
                        lblExTemp80.Text = NoDec(.extemp80) : lblExTemp60.Text = NoDec(.extemp60) : lblExTemp40.Text = NoDec(.extemp40)
                        lblHeatMain80.Text = NoDec(.mainheat80) : lblHeatMain60.Text = NoDec(.mainheat60) : lblHeatMain40.Text = NoDec(.mainheat40)
                        lblQExAvail80.Text = NoDec(.QExAvail80) : lblQExAvail60.Text = NoDec(.QExAvail60) : lblQExAvail40.Text = NoDec(.QExAvail40)
                        lblLTheat80.Text = NoDec(.lt_heat80) : lblLTheat60.Text = NoDec(.lt_heat60) : lblLTheat40.Text = NoDec(.lt_heat40)
                        lblOilCool100.Text = NoDec(.oilcool100) : lblOilCool80.Text = NoDec(.oilcool80) : lblOilCool60.Text = NoDec(.oilcool60) : lblOilCool40.Text = NoDec(.oilcool40)
                        lblJWin80.Text = NoDec(.jwin80) : lblJWin60.Text = NoDec(.jwin60) : lblJWin40.Text = NoDec(.jwin40)
                        lblJWout80.Text = .jw_out : lblJWout60.Text = .jw_out : lblJWout40.Text = .jw_out
                        lblJWFlowRate80.Text = .jw_flow : lblJWFlowRate60.Text = .jw_flow : lblJWFlowRate40.Text = .jw_flow
                        lblICin80.Text = .ic_in : lblICin60.Text = .ic_in : lblICin40.Text = .ic_in
                        lblICout80.Text = NoDec(.icout80) : lblICout60.Text = NoDec(.icout60) : lblICout40.Text = NoDec(.icout40)
                        lblICFlowRate80.Text = .ic_flow : lblICFlowRate60.Text = .ic_flow : lblICFlowRate40.Text = .ic_flow
                        lblFuelConHr80.Text = NoDec(.fuelcon80) : lblFuelConHr60.Text = NoDec(.fuelcon60) : lblFuelConHr40.Text = NoDec(.fuelcon40)
                        lblFuelConBhp80.Text = NoDec(.bHPhr80) : lblFuelConBhp60.Text = NoDec(.bHPhr60) : lblFuelConBhp40.Text = NoDec(.bHPhr40)
                        lblFuelKW80.Text = NoDec(.btuKWh80) : lblFuelKW60.Text = NoDec(.btuKWh60) : lblFuelKW40.Text = NoDec(.btuKWh40)
                        '============ RECOVERED HEAT PANEL ========================
                        lblJWtoPrimary80.Text = NoDec(.mainheat80) : lblJWtoPrimary60.Text = NoDec(.mainheat60) : lblJWtoPrimary40.Text = NoDec(.mainheat40)
                        lblExRecov80.Text = NoDec(.QEHRU80) : lblExRecov60.Text = NoDec(.QEHRU60) : lblExRecov40.Text = NoDec(.QEHRU40)
                        lblOiltoPrimary.Text = NoDec(.oilcool100) : lblOiltoPrimary80.Text = NoDec(.oilcool80) : lblOiltoPrimary60.Text = NoDec(.oilcool60) : lblOiltoPrimary40.Text = NoDec(.oilcool40)
                        .QJWRad80 *= -1 : lblJWRad80.Text = NoDec(.QJWRad80) : .QJWRad60 *= -1 : lblJWRad60.Text = NoDec(.QJWRad60) : .QJWRad40 *= -1 : lblJWRad40.Text = NoDec(.QJWRad40)
                        lblPrimHeat80.Text = NoDec((CInt(lblJWtoPrimary80.Text) + CInt(lblExRecov80.Text) + CInt(lblOiltoPrimary80.Text) + CInt(lblJWRad80.Text)))
                        lblPrimHeat60.Text = NoDec((CInt(lblJWtoPrimary60.Text) + CInt(lblExRecov60.Text) + CInt(lblOiltoPrimary60.Text) + CInt(lblJWRad60.Text)))
                        lblPrimHeat40.Text = NoDec((CInt(lblJWtoPrimary40.Text) + CInt(lblExRecov40.Text) + CInt(lblOiltoPrimary40.Text) + CInt(lblJWRad40.Text)))
                        lblQICHX80.Text = NoDec(.QICHX80) : lblQICHX60.Text = NoDec(.QICHX60) : lblQICHX40.Text = NoDec(.QICHX40)
                        .QICRad80 *= -1 : lblICRad80.Text = NoDec(.QICRad80) : .QICRad60 *= -1 : lblICRad60.Text = NoDec(.QICRad60) : .QICRad40 *= -1 : lblICRad40.Text = NoDec(.QICRad40)
                        lblSecHeat80.Text = NoDec(CDbl(lblQICHX80.Text) + CDbl(lblICRad80.Text))
                        lblSecHeat60.Text = NoDec(CDbl(lblQICHX60.Text) + CDbl(lblICRad60.Text))
                        lblSecHeat40.Text = NoDec(CDbl(lblQICHX40.Text) + CDbl(lblICRad40.Text))
                        lblEleEff80.Text = OneDec(.EleEff80) : lblEleEff60.Text = OneDec(.EleEff60) : lblEleEff40.Text = OneDec(.EleEff40)
                        lblThermEff80.Text = OneDec(.ThermEff80) : lblThermEff60.Text = OneDec(.ThermEff60) : lblThermEff40.Text = OneDec(.ThermEff40)
                        lblTotalEff80.Text = OneDec(.TotalEff80) : lblTotalEff60.Text = OneDec(.TotalEff60) : lblTotalEff40.Text = OneDec(.TotalEff40)
                        '============ PRIMARY CIRCUIT PANEL ========================
                        lblPWFlow80.Text = NoDec(.PwFlow80) : lblPWFlow60.Text = NoDec(.PwFlow60) : lblPWFlow40.Text = NoDec(.PwFlow40)
                        lblPwInAct80.Text = NoDec(.PwInActual80) : lblPwInAct60.Text = NoDec(.PwInActual60) : lblPwInAct40.Text = NoDec(.PwInActual40)
                        lblPwOutAct80.Text = NoDec(.PwOutActual80) : lblPwOutAct60.Text = NoDec(.PwOutActual60) : lblPwOutAct40.Text = NoDec(.PwOutActual40)
                        lblFluid2Type.Text = ._PrmCir_fluid.ToString : lblFluid2Percent.Text = String.Format("{0}%", ._f2pct)
                        '============ SECONDARY CIRCUIT PANEL ========================
                        lblSWFlow80.Text = NoDec(.SwFlow80) : lblSWFlow60.Text = NoDec(.SwFlow60) : lblSWFlow40.Text = NoDec(.SwFlow40)
                        lblSwInAct80.Text = NoDec(.SwInActual80) : lblSwInAct60.Text = NoDec(.SwInActual60) : lblSwInAct40.Text = NoDec(.SwInActual40)
                        lblSwOutAct80.Text = NoDec(.SwOutActual80) : lblSwOutAct60.Text = NoDec(.SwOutActual60) : lblSwOutAct40.Text = NoDec(.SwOutActual40)
                        lblFluid3Type.Text = ._SecCir_fluid.ToString : lblFluid3Percent.Text = String.Format("{0}%", ._f3pct)
                        '============ STEAM PANEL ========================
                        lblQsteam80.Text = NoDec(.QSteam80) : lblQsteam60.Text = NoDec(.QSteam60) : lblQsteam40.Text = NoDec(.QSteam40)
                        lblSteamProd80.Text = NoDec(.SteamProd80) : lblSteamProd60.Text = NoDec(.SteamProd60) : lblSteamProd40.Text = NoDec(.SteamProd40)
                End Select
            End With
        Catch ex As Exception : End Try ' DO NOTHING
    End Sub
#End Region
End Class