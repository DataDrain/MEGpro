' THIS CLASS HANDLES THE GENSET CREATION/CALCULATION DURING RUNTIME
Public Class Genset
    Private SQL As New SQLControl : Private _query As String = ""

#Region "DECLARATIONS"
#Region "--> Engine"
    ' STANDARD STATS
    Public avg As Integer ' <-- this is an integer to prevent query errors when looking for decimals
    Public jw_in As Integer : Public jw_out As Integer : Public jw_flow As Double
    Public ic_in As Integer : Public ic_out As Integer : Public ic_flow As Double
    Public exflow100 As Integer : Public exflow80 As Integer : Public exflow75 As Integer : Public exflow60 As Integer : Public exflow50 As Integer : Public exflow40 As Integer
    Public extemp100 As Integer : Public extemp80 As Integer : Public extemp75 As Integer : Public extemp60 As Integer : Public extemp50 As Integer : Public extemp40 As Integer
    Public mainheat100 As Integer : Public mainheat80 As Integer : Public mainheat75 As Integer : Public mainheat60 As Integer : Public mainheat50 As Integer : Public mainheat40 As Integer
    Public lt_heat100 As Integer : Public lt_heat80 As Integer : Public lt_heat75 As Integer : Public lt_heat60 As Integer : Public lt_heat50 As Integer : Public lt_heat40 As Integer
    Public fuelcon100 As Integer : Public fuelcon80 As Integer : Public fuelcon75 As Integer : Public fuelcon60 As Integer : Public fuelcon50 As Integer : Public fuelcon40 As Integer
    Public oilcool100 As Integer : Public oilcool80 As Integer : Public oilcool60 As Integer : Public oilcool40 As Integer
    Public elepow100 As Double : Public elepow80 As Double : Public elepow75 As Double : Public elepow60 As Double : Public elepow50 As Double : Public elepow40 As Double
    Public btuKWh As Integer : Public btuKWh80 As Integer : Public btuKWh75 As Integer : Public btuKWh60 As Integer : Public btuKW50 As Integer : Public btuKWh40 As Integer
    Public bHPhr As Integer : Public bHPhr80 As Integer : Public bHPhr75 As Integer : Public bHPhr60 As Integer : Public bHPhr50 As Integer : Public bHPhr40 As Integer
    Public engpow100 As Integer : Public engpow80 As Integer : Public engpow75 As Integer : Public engpow60 As Integer : Public engpow50 As Integer : Public engpow40 As Integer
    Public EleEff As Double : Public EleEff80 As Double : Public EleEff75 As Double : Public EleEff60 As Double : Public EleEff50 As Double : Public EleEff40 As Double
    Public ThermEff As Double : Public ThermEff80 As Double : Public ThermEff75 As Double : Public ThermEff60 As Double : Public ThermEff50 As Double : Public ThermEff40 As Double
    Public TotalEff As Double : Public TotalEff80 As Double : Public TotalEff75 As Double : Public TotalEff60 As Double : Public TotalEff50 As Double : Public TotalEff40 As Double
    ' MAN STATS
    Public _CoolHeat100u As Integer : Public _CoolHeat75u As Integer : Public _CoolHeat50u As Integer
    Public _MixHT100u As Integer : Public _MixHT75u As Integer : Public _MixHT50u As Integer
    Public _Radiation100u As Integer : Public _Radiation75u As Integer : Public _Radiation50u As Integer
#End Region
#Region "--> Generator"
    Public _genID As String : Public _genMFR As String : Public _genRPM As Integer : Public _genKW As Double : Public _genKVA As Double : Public _genVolts As Integer
    Public x6 As Double : Public x5 As Double : Public x4 As Double : Public x3 As Double : Public x2 As Double : Public x1 As Double : Public x0 As Double

    ' VARS TO FIND TRUE GEN EFFICIENCY
    Public genEff As Double : Public genLoad As Double
    Public KWeOut As Double : Public KWeOut100 As Double : Public KWeOut80 As Double : Public KWeOut75 As Double : Public KWeOut60 As Double : Public KWeOut50 As Double : Public KWeOut40 As Double
    Public loopCount As Integer = 0
#End Region
#Region "--> Calculation Vars"
    ' Q's
    Public QExAvail As Double : Public QExAvail80 As Double : Public QExAvail75 As Double : Public QExAvail60 As Double : Public QExAvail50 As Double : Public QExAvail40 As Double
    Public QEHRU As Double : Public QEHRU80 As Double : Public QEHRU75 As Double : Public QEHRU60 As Double : Public QEHRU50 As Double : Public QEHRU40 As Double
    Public QSteam As Double : Public QSteam80 As Double : Public QSteam75 As Double : Public QSteam60 As Double : Public QSteam50 As Double : Public QSteam40 As Double

    ' JW / IC
    Public jwout75 As Double : Public jwout50 As Double
    Public jwin80 As Double : Public jwin60 As Double : Public jwin40 As Double
    Public JWMassFlow As Double : Public JWdensity As Double : Public JWCp As Double
    Public ICMassFlow As Double : Public ICdensity As Double : Public ICcp As Double
    Public icout80 As Double : Public icout75 As Double : Public icout60 As Double : Public icout50 As Double : Public icout40 As Double

    ' CASE CALCULATIONS (PRIMARY)
    Public PwInActual As Double : Public PwInActual80 As Double : Public PwInActual75 As Double : Public PwInActual60 As Double : Public PwInActual50 As Double : Public PwInActual40 As Double
    Public PwOutActual As Double : Public PwOutActual80 As Double : Public PwOutActual75 As Double : Public PwOutActual60 As Double : Public PwOutActual50 As Double : Public PwOutActual40 As Double
    Public PWavg As Integer : Public Pwavg80 As Integer : Public PWavg75 As Integer : Public Pwavg60 As Integer : Public PWavg50 As Integer : Public Pwavg40 As Integer
    Public PwCp As Double : Public PwCp80 As Double : Public PwCp75 As Double : Public PwCp60 As Double : Public PwCp50 As Double : Public PwCp40 As Double
    Public PwDensity As Double : Public PwDensity80 As Double : Public PwDensity75 As Double : Public PwDensity60 As Double : Public PwDensity50 As Double : Public PwDensity40 As Double
    Public PostEHRU As Double : Public PostEHRU80 As Double : Public PostEHRU75 As Double : Public PostEHRU60 As Double : Public PostEHRU50 As Double : Public PostEHRU40 As Double
    Public PostHX As Double : Public PostHX80 As Double : Public PostHX75 As Double : Public PostHX60 As Double : Public PostHX50 As Double : Public PostHX40 As Double
    Public QHX As Integer : Public QHX80 As Integer : Public QHX75 As Integer : Public QHX60 As Integer : Public QHX50 As Integer : Public QHX40 As Integer
    Public QJWRad As Integer : Public QJWRad80 As Integer : Public QJWRad75 As Integer : Public QJWRad60 As Integer : Public QJWRad50 As Integer : Public QJWRad40 As Integer
    Public PwFlow As Integer : Public PwFlow80 As Integer : Public PwFlow75 As Integer : Public PwFlow60 As Integer : Public PwFlow50 As Integer : Public PwFlow40 As Integer

    ' CASE CALCULATIONS (SECONDARY)
    Public SwInActual As Double : Public SwInActual80 As Double : Public SwInActual75 As Double : Public SwInActual60 As Double : Public SwInActual50 As Double : Public SwInActual40 As Double
    Public SwOutActual As Double : Public SwOutActual80 As Double : Public SwOutActual75 As Double : Public SwOutActual60 As Double : Public SwOutActual50 As Double : Public SwOutActual40 As Double
    Public SWavg As Integer : Public Swavg80 As Integer : Public SWavg75 As Integer : Public Swavg60 As Integer : Public SWavg50 As Integer : Public SWavg40 As Integer
    Public SwCp As Double : Public SwCp80 As Double : Public SwCp75 As Double : Public SwCp60 As Double : Public SwCp50 As Double : Public SwCp40 As Double
    Public SwDensity As Double : Public SwDensity80 As Double : Public SwDensity75 As Double : Public SwDensity60 As Double : Public SwDensity50 As Double : Public SwDensity40 As Double
    ' DEREK...IS THERE A PostEHRU for secondary?
    Public PostICHX As Double : Public PostICHX80 As Double : Public PostICHX75 As Double : Public PostICHX60 As Double : Public PostICHX50 As Double : Public PostICHX40 As Double
    Public QICHX As Integer : Public QICHX80 As Integer : Public QICHX75 As Integer : Public QICHX60 As Integer : Public QICHX50 As Integer : Public QICHX40 As Integer
    Public QICRad As Integer : Public QICRad80 As Integer : Public QICRad75 As Integer : Public QICRad60 As Integer : Public QICRad50 As Integer : Public QICRad40 As Integer
    Public SWFlow As Integer : Public SwFlow80 As Integer : Public SwFlow75 As Integer : Public SwFlow60 As Integer : Public SwFlow50 As Integer : Public SwFlow40 As Integer

    ' STEAM
    Public SteamTemp As Double ' LOOK UP VALUE
    Public VaporEnth As Double ' LOOK UP VALUE
    Public SatLiq As Double ' LOOK UP VALUE
    Public SteamProduction As Integer : Public SteamProd80 As Integer : Public SteamProd75 As Integer : Public SteamProd60 As Integer : Public SteamProd50 As Integer : Public SteamProd40 As Integer

    ' PRIMARY HEAT
    Public QPrimary As Integer
    Public QLT As Integer ' <-- is this the same as QSecondary?

    Public CalcCase As Integer

    ' BOOLEAN CONDITIONS
    Public throwOut As Boolean = False : Public tooHot As Boolean = False
    Public FlagList As List(Of String)
    Public Enum TempFlags
        PW_In
        PW_Out
        SW_In
        SW_Out
    End Enum
#End Region
#Region "--> Constructor"
    Public _EngID As String : Public _MFR As String : Public _Model As String : Public _Fuel As String : Public _BurnType As String : Public _RPM As Integer : Public _PF As Decimal
    Private _user_MinExTemp As Integer : Public _user_StmPress As Integer : Private _user_Feed_H2O As Integer
    Public _user_PWin As Integer : Private _user_PWout As Integer : Private _user_SWin As Integer : Private _user_SWout As Integer
    Private _wantStm As Boolean : Private _wantEHRU As Boolean : Private _EHRUtoJW As Boolean : Private _EHRUtoPRM As Boolean
    Private _wantJW As Boolean : Private _wantLT As Boolean : Private _LTtoPRM As Boolean : Private _LTtoSEC As Boolean
    Private _EngCoolant_fluid As FluidType = Nothing : Private _f1pct As Integer
    Public _PrmCir_fluid As FluidType = Nothing : Public _f2pct As Integer
    Public _SecCir_fluid As FluidType = Nothing : Public _f3pct As Integer
    Private _OilToJW As Boolean : Private _OilToIC As Boolean
    ' CIRCUIT FLUIDS
    Public Enum FluidType
        Water
        Ethylene
        Propylene
        None
    End Enum
#End Region
#End Region

#Region "CONSTRUCTOR"
    ' SET INFORMATION FROM MAIN FORM
    Public Sub New(eID As String, mfr As String, minEx As Integer, stmPressure As Integer, feed As Integer, ppwIn As Integer, ppwOut As Integer, spwIn As Integer, spwOut As Integer, _
                   steam As Boolean, ehru As Boolean, eh2jw As Boolean, eh2primary As Boolean, jw As Boolean, lt As Boolean, lt2primary As Boolean, lt2second As Boolean, _
                   f1 As Integer, f2 As Integer, f3 As Integer, f1per As Integer, f2per As Integer, f3per As Integer, oil2jw As Boolean, oil2ic As Boolean)
        ' BASICS
        _EngID = eID : _MFR = mfr : _PF = frmMain.PowFactor
        ' USER INPUTS
        _user_MinExTemp = minEx : _user_StmPress = stmPressure : _user_Feed_H2O = feed : _user_PWin = ppwIn : _user_PWout = ppwOut : _user_SWin = spwIn : _user_SWout = spwOut
        ' BOOLEANS
        _wantStm = steam : _wantEHRU = ehru : _EHRUtoJW = eh2jw : _EHRUtoPRM = eh2primary : _wantJW = jw : _wantLT = lt : _LTtoPRM = lt2primary : _LTtoSEC = lt2second
        If _MFR = "Guascor" Then _OilToJW = oil2jw : _OilToIC = oil2ic
        ' FLUID TYPES
        _EngCoolant_fluid = f1 : _f1pct = f1per : _PrmCir_fluid = f2 : _f2pct = f2per : _SecCir_fluid = f3 : _f3pct = f3per
        ' ========    END IMPORTING FROM MAIN    ========

        FindCalcCase()

        GetEngineStats()

        GetGeneratorStats()

        GetGPMs()

        ' MANDITORY CALCS
        QExAvail = exflow100 * ExCp * (extemp100 - _user_MinExTemp)
        If _MFR = "Guascor" Then
            QExAvail80 = exflow80 * ExCp * (extemp80 - _user_MinExTemp) : QExAvail60 = exflow60 * ExCp * (extemp60 - _user_MinExTemp) : QExAvail40 = exflow40 * ExCp * (extemp40 - _user_MinExTemp)
        Else
            QExAvail75 = exflow75 * ExCp * (extemp75 - _user_MinExTemp) : QExAvail50 = exflow50 * ExCp * (extemp50 - _user_MinExTemp)
        End If

        ' CONDITIONAL CALCS
        If _wantStm = True Then CalcSteam()
        If _wantEHRU = True Then
            QEHRU = QExAvail - QSteam
            If _MFR = "Guascor" Then
                QEHRU80 = QExAvail80 - QSteam80 : QEHRU60 = QExAvail60 - QSteam60 : QEHRU40 = QExAvail40 - QSteam40
            Else
                QEHRU75 = QExAvail75 - QSteam75 : QEHRU50 = QExAvail50 - QSteam50
            End If
        End If

        CaseCalculations()

        EfficiencyAndFuelConCalcs()
    End Sub
#End Region

#Region "QUERY SUBS"
#Region "--> Engine Stats"
    Public Sub GetEngineStats()
        Select Case _MFR
            Case "Guascor"
                SQL.ExecQuery(String.Format("SELECT {0} FROM Engines WHERE id='{1}'", GUASCOR_GENSET, _EngID))
                _Model = _get(SQL.DBDS, "model") : _RPM = _get(SQL.DBDS, "rpm") : _Fuel = _get(SQL.DBDS, "fuel") : elepow100 = _get(SQL.DBDS, "elepow100")
                elepow80 = _get(SQL.DBDS, "elepow80") : elepow60 = _get(SQL.DBDS, "elepow60") : elepow80 = _get(SQL.DBDS, "elepow60")
                engpow100 = _get(SQL.DBDS, "engpow100") : engpow80 = _get(SQL.DBDS, "engpow80") : engpow60 = _get(SQL.DBDS, "engpow60") : engpow40 = _get(SQL.DBDS, "engpow40")
                jw_out = _get(SQL.DBDS, "jw_out") : jw_flow = _get(SQL.DBDS, "jw_flow") : ic_out = _get(SQL.DBDS, "ic_out") : ic_flow = _get(SQL.DBDS, "ic_flow")
                exflow100 = _get(SQL.DBDS, "exflow100") : exflow80 = _get(SQL.DBDS, "exflow80") : exflow60 = _get(SQL.DBDS, "exflow60") : exflow40 = _get(SQL.DBDS, "exflow40")
                extemp100 = _get(SQL.DBDS, "extemp100") : extemp80 = _get(SQL.DBDS, "extemp80") : extemp60 = _get(SQL.DBDS, "extemp60") : extemp40 = _get(SQL.DBDS, "extemp40")
                fuelcon100 = _get(SQL.DBDS, "fuelcon100_u") : fuelcon80 = _get(SQL.DBDS, "fuelcon80_u") : fuelcon60 = _get(SQL.DBDS, "fuelcon60_u") : fuelcon40 = _get(SQL.DBDS, "fuelcon40_u")
                mainheat100 = _get(SQL.DBDS, "mainheat100_u") : mainheat80 = _get(SQL.DBDS, "mainheat80_u") : mainheat60 = _get(SQL.DBDS, "mainheat60_u") : mainheat40 = _get(SQL.DBDS, "mainheat40_u")
                lt_heat100 = _get(SQL.DBDS, "lt_heat100_u") : lt_heat80 = _get(SQL.DBDS, "lt_heat80_u") : lt_heat60 = _get(SQL.DBDS, "lt_heat60_u") : lt_heat40 = _get(SQL.DBDS, "lt_heat40_u")
                oilcool100 = _get(SQL.DBDS, "oil_cooler100_u") : oilcool80 = _get(SQL.DBDS, "oil_cooler80_u") : oilcool60 = _get(SQL.DBDS, "oil_cooler60_u") : oilcool40 = _get(SQL.DBDS, "oil_cooler40_u")
                ' GET JW_IN & IC_OUT
                If _OilToJW Then
                    jw_in = UberLoop("guascorInlet", jw_out, _EngCoolant_fluid, (mainheat100 + oilcool100), jw_flow)
                    ic_in = UberLoop("guascorOutlet", ic_out, _EngCoolant_fluid, lt_heat100, ic_flow)
                Else
                    jw_in = UberLoop("guascorInlet", jw_out, _EngCoolant_fluid, mainheat100, jw_flow)
                    ic_in = UberLoop("guascorOutlet", ic_out, _EngCoolant_fluid, (lt_heat100 + oilcool100), ic_flow)
                End If
            Case "MTU"
                SQL.ExecQuery(String.Format("SELECT {0} FROM Engines WHERE id='{1}'", MTU_GENSET, _EngID))
                _Model = _get(SQL.DBDS, "model") : _RPM = _get(SQL.DBDS, "rpm") : _Fuel = _get(SQL.DBDS, "fuel")
                KWeOut100 = _get(SQL.DBDS, "elepow100") : KWeOut75 = _get(SQL.DBDS, "elepow75") : KWeOut50 = _get(SQL.DBDS, "elepow50") : _genVolts = _get(SQL.DBDS, "voltage")
                engpow100 = _get(SQL.DBDS, "engpow100") : engpow75 = _get(SQL.DBDS, "engpow75") : engpow50 = _get(SQL.DBDS, "engpow50")
                jw_in = _get(SQL.DBDS, "jw_in") : jw_out = _get(SQL.DBDS, "jw_out") : ic_in = _get(SQL.DBDS, "ic_in") : ic_out = _get(SQL.DBDS, "ic_out")
                exflow100 = _get(SQL.DBDS, "exflow100") : exflow75 = _get(SQL.DBDS, "exflow75") : exflow50 = _get(SQL.DBDS, "exflow50")
                extemp100 = _get(SQL.DBDS, "extemp100") : extemp75 = _get(SQL.DBDS, "extemp75") : extemp50 = _get(SQL.DBDS, "extemp50")
                fuelcon100 = _get(SQL.DBDS, "fuelcon100_u") : fuelcon75 = _get(SQL.DBDS, "fuelcon75_u") : fuelcon50 = _get(SQL.DBDS, "fuelcon50_u")
                mainheat100 = _get(SQL.DBDS, "mainheat100_u") : mainheat75 = _get(SQL.DBDS, "mainheat75_u") : mainheat50 = _get(SQL.DBDS, "mainheat50_u")
                lt_heat100 = _get(SQL.DBDS, "lt_heat100_u") : lt_heat75 = _get(SQL.DBDS, "lt_heat75_u") : lt_heat50 = _get(SQL.DBDS, "lt_heat50_u")
        End Select
        If Not String.IsNullOrEmpty(SQL.Exception) Then MsgBox(SQL.Exception)
    End Sub
#End Region
#Region "--> Generator Stats"
    Public Sub GetGeneratorStats()
        If _MFR = "MTU" Then _genID = "Included" : Exit Sub
        ' PAIR GENERATOR TO ENGINE
        SQL.AddParam("@id", _EngID)
        SQL.ExecQuery("SELECT TOP(1) g.id FROM generators AS g, engines AS e WHERE e.id=@id AND e.rpm = g.rpm AND g.elepow > e.elepow100")
        _genID = _get(SQL.DBDS, "id")
        ' GET GENERATOR STATS ACCORDING TO PAIRED ID
        SQL.AddParam("@gen", _genID)
        SQL.ExecQuery("SELECT * FROM Generators WHERE id=@gen")
        If String.IsNullOrEmpty(SQL.Exception) Then
            _genMFR = SQL.DBDS.Tables(0).Rows(0)("mfr")
            _genKW = SQL.DBDS.Tables(0).Rows(0)("elepow")
            _genVolts = SQL.DBDS.Tables(0).Rows(0)("voltage")
            _genKVA = SQL.DBDS.Tables(0).Rows(0)("kva")
            Select Case _PF
                Case 0.8
                    x6 = _get(SQL.DBDS, "p8x6") : x5 = _get(SQL.DBDS, "p8x5") : x4 = _get(SQL.DBDS, "p8x4") : x3 = _get(SQL.DBDS, "p8x3")
                    x2 = _get(SQL.DBDS, "p8x2") : x1 = _get(SQL.DBDS, "p8x1") : x0 = _get(SQL.DBDS, "p8x0")
                Case 0.9
                    x6 = _get(SQL.DBDS, "p9x6") : x5 = _get(SQL.DBDS, "p9x5") : x4 = _get(SQL.DBDS, "p9x4") : x3 = _get(SQL.DBDS, "p9x3")
                    x2 = _get(SQL.DBDS, "p9x2") : x1 = _get(SQL.DBDS, "p9x1") : x0 = _get(SQL.DBDS, "p9x0")
                Case 1.0
                    x6 = _get(SQL.DBDS, "p1x6") : x5 = _get(SQL.DBDS, "p1x5") : x4 = _get(SQL.DBDS, "p1x4") : x3 = _get(SQL.DBDS, "p1x3")
                    x2 = _get(SQL.DBDS, "p1x2") : x1 = _get(SQL.DBDS, "p1x1") : x0 = _get(SQL.DBDS, "p1x0")
            End Select
        Else
            MsgBox(SQL.Exception)
        End If
        GenEfficiency()
    End Sub
    Private Sub GenEfficiency()
        If _MFR <> "MTU" Then
            KWeOut100 = CalcGenEff(engpow100)
            If _MFR = "Guascor" Then
                KWeOut80 = CalcGenEff(engpow80) : KWeOut60 = CalcGenEff(engpow60) : KWeOut40 = CalcGenEff(engpow40)
            Else
                KWeOut75 = CalcGenEff(engpow75) : KWeOut50 = CalcGenEff(engpow50)
            End If
        End If
    End Sub
    Private Function CalcGenEff(bhp As Double) As Double
        Dim kwEtest As Double = bhp * ELE_CONVERSION * 0.95 : Dim i As Integer
        genLoad = kwEtest / (_PF * _genKVA)
        genEff = ((genLoad ^ 6 * x6) + (genLoad ^ 5 * x5) + (genLoad ^ 4 * x4) + (genLoad ^ 3 * x3) + (genLoad ^ 2 * x2) + (genLoad ^ 1 * x1) + (x0)) / 100
        KWeOut = bhp * ELE_CONVERSION * genEff
        ' ENTER LOOP
        While (i < 5)
            If System.Math.Abs(kwEtest - KWeOut) <= 0.5 Then Exit While
            kwEtest = KWeOut
            genLoad = kwEtest / (_PF * _genKVA)
            genEff = ((genLoad ^ 6 * x6) + (genLoad ^ 5 * x5) + (genLoad ^ 4 * x4) + (genLoad ^ 3 * x3) + (genLoad ^ 2 * x2) + (genLoad ^ 1 * x1) + (x0)) / 100
            KWeOut = bhp * ELE_CONVERSION * genEff
            i += 1
        End While
        Return KWeOut
    End Function
#End Region
#Region "--> Circuit Fluids"
    Public Function GetFluidValue(type As String, Temp As Double, Cp As Boolean, Density As Boolean, Optional Percent As Double = Nothing) As Double
        Dim tblName As String = "" : Dim tblPrefix As String = "" : Dim tblSuffix As String = ""
        Dim colName As String = ""
        ' PREPARE THE QUERY BASED UPON PARAMETERS
        If type = "Ethylene" Or type = "Propylene" Then
            tblPrefix = type
            If Cp = True Then tblSuffix = "Cp"
            If Density = True Then tblSuffix = "Density"
            tblName = tblPrefix & tblSuffix
            colName = "p" & Percent
            _query = String.Format("SELECT TOP(1) {0} FROM {1} WHERE Temp>={2}", colName, tblName, Temp)
        ElseIf type = "Water" Then
            tblPrefix = type : tblSuffix = "Properties"
            tblName = tblPrefix & tblSuffix
            If Cp = True Then colName = "Cp"
            If Density = True Then colName = "Density"
        Else : Return 0 : End If

        _query = String.Format("SELECT TOP(1) {0} FROM {1} WHERE Temp>={2}", colName, tblName, Temp)
        SQL.ExecQuery(_query)

        If String.IsNullOrEmpty(SQL.Exception) Then
            If Not String.IsNullOrEmpty(SQL.DBDS.Tables(0).Rows(0)(colName)) Then
                Return SQL.DBDS.Tables(0).Rows(0)(colName)
            Else
                MsgBox("Null value found during lookup")
            End If
        Else
            MsgBox(SQL.Exception) : End If
        Return 0
    End Function
#End Region
#End Region

#Region "CALCULATION SUBS"
#Region "--> Calc Cases"
    Public Sub FindCalcCase()
        If _wantLT = False Then
            If _EHRUtoJW = True Or _wantEHRU = False Then
                CalcCase = 1
            Else
                CalcCase = 4
            End If
        ElseIf _LTtoSEC = True Then
            If _EHRUtoJW = True Or _wantEHRU = False Then
                CalcCase = 2
            Else
                CalcCase = 5
            End If
        ElseIf _LTtoPRM = True Then
            If _EHRUtoJW = True Or _wantEHRU = False Then
                CalcCase = 3
            Else
                CalcCase = 6
            End If
        End If
    End Sub
    '/////////////////////////////////////////////////////////
    Public Sub CaseCalculations()
        Select Case CalcCase
            Case 1
                DeterminePWtemps()
                CalcPWflow()
                QLT = 0
                QPrimary = QHX
            Case 2
                DeterminePWtemps()
                CalcPWflow()
                DetermineSWtemps()
                CalcSWflow()
                QLT = QICHX
                QPrimary = QHX
            Case 3
            Case 4
            Case 5
            Case 6
        End Select
    End Sub
#End Region
#Region "--> Primary PW"
    Public Sub DeterminePWtemps()
        If _MFR = "Guascor" Then ' inlet will vary, JWout is constant
            If _OilToJW Then
                jwin80 = UberLoop("inlet", jw_out, _EngCoolant_fluid, (mainheat80 + oilcool80), JWMassFlow) : jwin60 = UberLoop("inlet", jw_out, _EngCoolant_fluid, (mainheat60 + oilcool60), JWMassFlow) : jwin40 = UberLoop("inlet", jw_out, _EngCoolant_fluid, (mainheat40 + oilcool40), JWMassFlow)
            Else
                jwin80 = UberLoop("inlet", jw_out, _EngCoolant_fluid, mainheat80, JWMassFlow) : jwin60 = UberLoop("inlet", jw_out, _EngCoolant_fluid, mainheat60, JWMassFlow) : jwin40 = UberLoop("inlet", jw_out, _EngCoolant_fluid, mainheat40, JWMassFlow)
            End If
            ' DETERMINE IF PRIMARY PW TEMPS ARE ACCEPTABLE
            If Not JWMassFlow = 0 Or Not JWCp = 0 Then
                PostEHRU = UberLoop("outlet", jw_out, _EngCoolant_fluid, QEHRU, JWMassFlow)
                PostEHRU80 = UberLoop("outlet", jw_out, _EngCoolant_fluid, QEHRU80, JWMassFlow) : PostEHRU60 = UberLoop("outlet", jw_out, _EngCoolant_fluid, QEHRU60, JWMassFlow) : PostEHRU40 = UberLoop("outlet", jw_out, _EngCoolant_fluid, QEHRU40, JWMassFlow)
            Else
                PostEHRU = 0 : PostEHRU80 = 0 : PostEHRU60 = 0 : PostEHRU40 = 0
            End If
            ' PRIMARY OUT ACTUALS
            If PostEHRU - _user_PWout < MinApproachTemp Then PwOutActual = PostEHRU - MinApproachTemp : tooHot = True Else PwOutActual = _user_PWout
            If PostEHRU80 - _user_PWout < MinApproachTemp Then PwOutActual80 = PostEHRU80 - MinApproachTemp : tooHot = True Else PwOutActual80 = _user_PWout
            If PostEHRU60 - _user_PWout < MinApproachTemp Then PwOutActual60 = PostEHRU60 - MinApproachTemp : tooHot = True Else PwOutActual60 = _user_PWout
            If PostEHRU40 - _user_PWout < MinApproachTemp Then PwOutActual40 = PostEHRU40 - MinApproachTemp : tooHot = True Else PwOutActual40 = _user_PWout
            ' PRIMARY IN ACTUALS
            If PostEHRU - _user_PWin < MinApproachTemp Then PwInActual = PostEHRU - MinApproachTemp : tooHot = True Else PwInActual = _user_PWin
            If PostEHRU80 - _user_PWin < MinApproachTemp Then PwInActual80 = PostEHRU80 - MinApproachTemp : tooHot = True Else PwInActual80 = _user_PWin
            If PostEHRU60 - _user_PWin < MinApproachTemp Then PwInActual60 = PostEHRU60 - MinApproachTemp : tooHot = True Else PwInActual60 = _user_PWin
            If PostEHRU40 - _user_PWin < MinApproachTemp Then PwInActual40 = PostEHRU40 - MinApproachTemp : tooHot = True Else PwInActual40 = _user_PWin
            ' POST_HX
            If jw_in - _user_PWin < MinApproachTemp Then PostHX = PwInActual + MinApproachTemp Else PostHX = jw_in
            If jwin80 - _user_PWin < MinApproachTemp Then PostHX80 = PwInActual80 + MinApproachTemp Else PostHX80 = jwin80
            If jwin60 - _user_PWin < MinApproachTemp Then PostHX60 = PwInActual60 + MinApproachTemp Else PostHX60 = jwin60
            If jwin40 - _user_PWin < MinApproachTemp Then PostHX40 = PwInActual40 + MinApproachTemp Else PostHX40 = jwin40
        Else '==== 25's ====
            jwout75 = UberLoop("outlet", jw_in, _EngCoolant_fluid, mainheat75, JWMassFlow) : jwout50 = UberLoop("outlet", jw_in, _EngCoolant_fluid, mainheat50, JWMassFlow)
            ' DETERMINES IF PRIMARY PW TEMPS ARE ACCEPTABLE
            If Not JWMassFlow = 0 Or Not JWCp = 0 Then
                PostEHRU = UberLoop("outlet", jw_out, _EngCoolant_fluid, QEHRU, JWMassFlow) ' <-- PostEHRU = QEHRU / (JwMassFlow * JwCp) + JwOut
                PostEHRU75 = UberLoop("outlet", jwout75, _EngCoolant_fluid, QEHRU75, JWMassFlow) : PostEHRU50 = UberLoop("outlet", jwout50, _EngCoolant_fluid, QEHRU50, JWMassFlow)
            Else
                PostEHRU = 0 : PostEHRU75 = 0 : PostEHRU50 = 0 ' add 20's later
            End If
            ' PRIMARY OUT ACTUALS
            If PostEHRU - _user_PWout < MinApproachTemp Then PwOutActual = PostEHRU - MinApproachTemp : tooHot = True Else PwOutActual = _user_PWout
            If PostEHRU75 - _user_PWout < MinApproachTemp Then PwOutActual75 = PostEHRU75 - MinApproachTemp : tooHot = True Else PwOutActual75 = _user_PWout
            If PostEHRU50 - _user_PWout < MinApproachTemp Then PwOutActual50 = PostEHRU50 - MinApproachTemp : tooHot = True Else PwOutActual50 = _user_PWout
            ' PRIMARY IN ACTUALS
            If PostEHRU - _user_PWin < MinApproachTemp Then PwInActual = PostEHRU - MinApproachTemp : tooHot = True Else PwInActual = _user_PWin
            If PostEHRU75 - _user_PWin < MinApproachTemp Then PwInActual75 = PostEHRU75 - MinApproachTemp : tooHot = True Else PwInActual75 = _user_PWin
            If PostEHRU50 - _user_PWin < MinApproachTemp Then PwInActual50 = PostEHRU50 - MinApproachTemp : tooHot = True Else PwInActual50 = _user_PWin
            ' POST_HX
            If jw_in - _user_PWin < MinApproachTemp Then
                PostHX = PwInActual + MinApproachTemp : PostHX75 = PwInActual75 + MinApproachTemp : PostHX50 = PwInActual50 + MinApproachTemp
            Else
                PostHX = jw_in : PostHX75 = jw_in : PostHX50 = jw_in
            End If
        End If
    End Sub
    Public Sub CalcPWflow() ' CALC HEAT TRANSFERRED THROUGH HEAT EXCHANGER & HEAT LOST TO RADIATOR
        ' SOLVE FOR QHX
        QHX = UberLoop("Q", PostEHRU, _EngCoolant_fluid, PostHX, JWMassFlow) ' <--- QHX = JWMassFlow * JWCp * (PostEHRU - PostHX)
        If _MFR = "Guascor" Then
            QHX80 = UberLoop("Q", PostEHRU80, _EngCoolant_fluid, PostHX80, JWMassFlow) : QHX60 = UberLoop("Q", PostEHRU60, _EngCoolant_fluid, PostHX60, JWMassFlow) : QHX40 = UberLoop("Q", PostEHRU40, _EngCoolant_fluid, PostHX40, JWMassFlow)
        Else
            QHX75 = UberLoop("Q", PostEHRU75, _EngCoolant_fluid, PostHX75, JWMassFlow) : QHX50 = UberLoop("Q", PostEHRU50, _EngCoolant_fluid, PostHX50, JWMassFlow)
        End If
        ' SOLVE FOR QJWRAD
        QJWRad = UberLoop("Q", PostHX, _EngCoolant_fluid, jw_in, JWMassFlow) ' <--- 'QJWRad = JWMassFlow * JWCp * (PostHX - JWin)
        If _MFR = "Guascor" Then
            QJWRad80 = UberLoop("Q", PostHX80, _EngCoolant_fluid, jwin80, JWMassFlow) : QJWRad60 = UberLoop("Q", PostHX60, _EngCoolant_fluid, jwin60, JWMassFlow) : QJWRad40 = UberLoop("Q", PostHX40, _EngCoolant_fluid, jwin40, JWMassFlow)
        Else
            QJWRad75 = UberLoop("Q", PostHX75, _EngCoolant_fluid, jw_in, JWMassFlow) : QJWRad50 = UberLoop("Q", PostHX50, _EngCoolant_fluid, jw_in, JWMassFlow)
        End If
        ' PW FLOW RATE
        getPrimaryFluidVals()
        If PwInActual <= 0 Or PwOutActual <= 0 Then
            PwFlow = QHX / (PwDensity * PwCp * ConversionRatio)
            If _MFR = "Guascor" Then
                PwFlow80 = QHX80 / (PwDensity80 * PwCp80 * ConversionRatio) : PwFlow60 = QHX60 / (PwDensity60 * PwCp60 * ConversionRatio) : PwFlow40 = QHX40 / (PwDensity40 * PwCp40 * ConversionRatio)
            Else
                PwFlow75 = QHX75 / (PwDensity75 * PwCp75 * ConversionRatio) : PwFlow50 = QHX50 / (PwDensity50 * PwCp50 * ConversionRatio)
            End If
        Else
            PwFlow = QHX / (PwDensity * PwCp * (PwOutActual - PwInActual) * ConversionRatio)
            If _MFR = "Guascor" Then
                PwFlow80 = QHX80 / (PwDensity80 * PwCp80 * (PwOutActual80 - PwInActual80) * ConversionRatio) : PwFlow60 = QHX60 / (PwDensity60 * PwCp60 * (PwOutActual60 - PwInActual60) * ConversionRatio) : PwFlow40 = QHX40 / (PwDensity40 * PwCp40 * (PwOutActual40 - PwInActual40) * ConversionRatio)
            Else
                PwFlow75 = QHX75 / (PwDensity75 * PwCp75 * (PwOutActual75 - PwInActual75) * ConversionRatio) : PwFlow50 = QHX50 / (PwDensity50 * PwCp50 * (PwOutActual50 - PwInActual50) * ConversionRatio)
            End If
        End If
    End Sub
    Public Sub getPrimaryFluidVals()
        ' GET AVERAGES TO PLUG INTO QUERIES
        PWavg = (PwInActual + PwOutActual) / 2
        If _MFR = "Guascor" Then
            Pwavg80 = (PwInActual80 + PwOutActual80) / 2 : Pwavg60 = (PwInActual60 + PwOutActual60) / 2 : Pwavg40 = (PwInActual40 + PwOutActual40) / 2
        Else
            PWavg75 = (PwInActual75 + PwOutActual75) / 2 : PWavg50 = (PwInActual50 + PwOutActual50) / 2
        End If
        ' FIND JW_CP
        If _PrmCir_fluid = FluidType.Water Then
            PwCp = GetFluidValue(_PrmCir_fluid.ToString, PWavg, True, False)
            If _MFR = "Guascor" Then
                PwCp80 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg80, True, False) : PwCp60 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg60, True, False) : PwCp40 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg40, True, False)
            Else
                PwCp75 = GetFluidValue(_PrmCir_fluid.ToString, PWavg75, True, False) : PwCp50 = GetFluidValue(_PrmCir_fluid.ToString, PWavg50, True, False)
            End If
        Else
            PwCp = GetFluidValue(_PrmCir_fluid.ToString, PWavg, True, False, _f2pct)
            If _MFR = "Guascor" Then
                PwCp80 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg80, True, False, _f2pct) : PwCp60 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg60, True, False, _f2pct) : PwCp40 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg40, True, False, _f2pct)
            Else
                PwCp75 = GetFluidValue(_PrmCir_fluid.ToString, PWavg75, True, False, _f2pct) : PwCp50 = GetFluidValue(_PrmCir_fluid.ToString, PWavg50, True, False, _f2pct)
            End If
        End If
        ' FIND JW_DENSITY
        If _PrmCir_fluid = FluidType.Water Then
            PwDensity = GetFluidValue(_PrmCir_fluid.ToString, PWavg, False, True)
            If _MFR = "Guascor" Then
                PwDensity80 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg80, False, True) : PwDensity60 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg60, False, True) : PwDensity40 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg40, False, True)
            Else
                PwDensity75 = GetFluidValue(_PrmCir_fluid.ToString, PWavg75, False, True) : PwDensity50 = GetFluidValue(_PrmCir_fluid.ToString, PWavg50, False, True)
            End If
        Else
            PwDensity = GetFluidValue(_PrmCir_fluid.ToString, PWavg, False, True, _f2pct)
            If _MFR = "Guascor" Then
                PwDensity80 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg80, False, True, _f2pct) : PwDensity60 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg60, False, True, _f2pct) : PwDensity40 = GetFluidValue(_PrmCir_fluid.ToString, Pwavg40, False, True, _f2pct)
            Else
                PwDensity75 = GetFluidValue(_PrmCir_fluid.ToString, PWavg75, False, True, _f2pct) : PwDensity50 = GetFluidValue(_PrmCir_fluid.ToString, PWavg50, False, True, _f2pct)
            End If
        End If
    End Sub
#End Region
#Region "--> 2nd PW"
    Public Sub DetermineSWtemps()
        If _MFR = "Guascor" Then
            icout80 = UberLoop("outlet", ic_in, _SecCir_fluid, lt_heat80, ICMassFlow) : icout60 = UberLoop("outlet", ic_in, _SecCir_fluid, lt_heat60, ICMassFlow) : icout40 = UberLoop("outlet", ic_in, _SecCir_fluid, lt_heat40, ICMassFlow)
        Else
            icout75 = UberLoop("outlet", ic_in, _EngCoolant_fluid, lt_heat75, ICMassFlow) : icout50 = UberLoop("outlet", ic_in, _EngCoolant_fluid, lt_heat50, ICMassFlow)
        End If
        ' SW_OUT ACTUALS
        If ic_out - _user_SWout < MinApproachTemp Then SwOutActual = ic_out - MinApproachTemp : tooHot = True Else SwOutActual = _user_SWout
        If _MFR = "Guascor" Then
            If icout80 - _user_SWout < MinApproachTemp Then SwOutActual80 = icout80 - MinApproachTemp : tooHot = True Else SwOutActual80 = _user_SWout
            If icout60 - _user_SWout < MinApproachTemp Then SwOutActual60 = icout60 - MinApproachTemp : tooHot = True Else SwOutActual60 = _user_SWout
            If icout40 - _user_SWout < MinApproachTemp Then SwOutActual40 = icout40 - MinApproachTemp : tooHot = True Else SwOutActual40 = _user_SWout
        Else
            If icout75 - _user_SWout < MinApproachTemp Then SwOutActual75 = icout75 - MinApproachTemp : tooHot = True Else SwOutActual75 = _user_SWout
            If icout50 - _user_SWout < MinApproachTemp Then SwOutActual50 = icout50 - MinApproachTemp : tooHot = True Else SwOutActual50 = _user_SWout
        End If
        ' SW_IN ACTUALS
        If ic_out - _user_SWin < MinApproachTemp Then SwInActual = ic_out - MinApproachTemp : tooHot = True Else SwInActual = _user_SWin
        If _MFR = "Guascor" Then
            If icout80 - _user_SWin < MinApproachTemp Then SwInActual80 = icout80 - MinApproachTemp : tooHot = True Else SwInActual80 = _user_SWin
            If icout60 - _user_SWin < MinApproachTemp Then SwInActual60 = icout60 - MinApproachTemp : tooHot = True Else SwInActual60 = _user_SWin
            If icout40 - _user_SWin < MinApproachTemp Then SwInActual40 = icout40 - MinApproachTemp : tooHot = True Else SwInActual40 = _user_SWin
        Else
            If icout75 - _user_SWin < MinApproachTemp Then SwInActual75 = icout75 - MinApproachTemp : tooHot = True Else SwInActual75 = _user_SWin
            If icout50 - _user_SWin < MinApproachTemp Then SwInActual50 = icout50 - MinApproachTemp : tooHot = True Else SwInActual50 = _user_SWin
        End If
        ' POST IC_HX
        If ic_in - _user_SWin < MinApproachTemp Then
            PostICHX = SwInActual + MinApproachTemp
            If _MFR = "Guascor" Then
                PostICHX80 = SwInActual80 + MinApproachTemp : PostICHX60 = SwInActual60 + MinApproachTemp : PostICHX40 = SwInActual40 + MinApproachTemp
            Else
                PostICHX75 = SwInActual75 + MinApproachTemp : PostICHX50 = SwInActual50 + MinApproachTemp
            End If
            If PostICHX >= ic_out Then PostICHX = ic_out
            If _MFR = "Guascor" Then
                If PostICHX80 >= icout80 Then PostICHX80 = icout80 : If PostICHX60 >= icout80 Then PostICHX60 = icout60 : If PostICHX40 >= icout80 Then PostICHX40 = icout40
            Else
                If PostICHX75 >= icout75 Then PostICHX75 = icout75 : If PostICHX50 >= icout50 Then PostICHX50 = icout50
            End If
        Else
            PostICHX = ic_in
            If _MFR = "Guascor" Then
                PostICHX80 = ic_in : PostICHX60 = ic_in : PostICHX40 = ic_in
            Else
                PostICHX75 = ic_in : PostICHX50 = ic_in
            End If
        End If
    End Sub

    Public Sub CalcSWflow() ' CALC HEAT TRANSFERRED THROUGH HEAT EXCHANGER & HEAT LOST TO RADIATOR
        ' QICHX
        QICHX = UberLoop("Q", PostICHX, _SecCir_fluid, ic_out, ICMassFlow)
        If _MFR = "Guascor" Then
            QICHX80 = UberLoop("Q", PostICHX80, _SecCir_fluid, icout80, ICMassFlow) : QICHX60 = UberLoop("Q", PostICHX60, _SecCir_fluid, icout60, ICMassFlow) : QICHX40 = UberLoop("Q", PostICHX40, _SecCir_fluid, icout40, ICMassFlow)
        Else
            QICHX75 = UberLoop("Q", PostICHX75, _SecCir_fluid, icout75, ICMassFlow) : QICHX50 = UberLoop("Q", PostICHX50, _SecCir_fluid, icout50, ICMassFlow)
        End If
        'QICRAD
        QICRad = UberLoop("Q", PostICHX, _SecCir_fluid, ic_in, ICMassFlow)
        If _MFR = "Guascor" Then
            QICRad80 = UberLoop("Q", PostICHX80, _SecCir_fluid, ic_in, ICMassFlow) : QICRad60 = UberLoop("Q", PostICHX60, _SecCir_fluid, ic_in, ICMassFlow) : QICRad40 = UberLoop("Q", PostICHX40, _SecCir_fluid, ic_in, ICMassFlow)
        Else
            QICRad75 = UberLoop("Q", PostICHX75, _SecCir_fluid, ic_in, ICMassFlow) : QICRad50 = UberLoop("Q", PostICHX50, _SecCir_fluid, ic_in, ICMassFlow)
        End If

        ' CALC 2ndary FLOW RATE (IC FLOW)
        getSecondaryFluidVals()
        If SwOutActual = SwInActual Then SWFlow = 0 Else SWFlow = QICHX / (SwDensity * SwCp * (SwOutActual - SwInActual) * ConversionRatio)
        If _MFR = "Guascor" Then
            If SwOutActual80 = SwInActual80 Then SwFlow80 = 0 Else SwFlow80 = QICHX80 / (SwDensity80 * SwCp80 * (SwOutActual80 - SwInActual80) * ConversionRatio)
            If SwOutActual60 = SwInActual60 Then SwFlow60 = 0 Else SwFlow60 = QICHX60 / (SwDensity60 * SwCp60 * (SwOutActual60 - SwInActual60) * ConversionRatio)
            If SwOutActual40 = SwInActual40 Then SwFlow40 = 0 Else SwFlow40 = QICHX40 / (SwDensity40 * SwCp40 * (SwOutActual40 - SwInActual40) * ConversionRatio)
        Else
            If SwOutActual75 = SwInActual75 Then SwFlow75 = 0 Else SwFlow75 = QICHX75 / (SwDensity75 * SwCp75 * (SwOutActual75 - SwInActual75) * ConversionRatio)
            If SwOutActual50 = SwInActual50 Then SwFlow50 = 0 Else SwFlow50 = QICHX50 / (SwDensity50 * SwCp50 * (SwOutActual50 - SwInActual50) * ConversionRatio)
        End If
    End Sub
    Public Sub getSecondaryFluidVals()
        SWavg = (SwInActual + SwOutActual) / 2
        If _MFR = "Guascor" Then
            Swavg80 = (SwInActual80 + SwOutActual80) / 2 : Swavg60 = (SwInActual60 + SwOutActual60) / 2 : SWavg40 = (SwInActual40 + SwOutActual40) / 2
        Else
            SWavg75 = (SwInActual75 + SwOutActual75) / 2 : SWavg50 = (SwInActual50 + SwOutActual50) / 2
        End If

        ' FIND SW_CP
        If _SecCir_fluid = FluidType.Water Then
            If _MFR = "Guascor" Then
                If SWavg = Swavg80 AndAlso Swavg80 = Swavg60 AndAlso Swavg60 = SWavg40 Then
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False) : SwCp80 = SwCp : SwCp60 = SwCp : SwCp40 = SwCp
                Else
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False)
                    SwCp80 = GetFluidValue(_SecCir_fluid.ToString, Swavg80, True, False)
                    SwCp60 = GetFluidValue(_SecCir_fluid.ToString, Swavg60, True, False)
                    SwCp40 = GetFluidValue(_SecCir_fluid.ToString, SWavg40, True, False)
                End If
            Else ' IF THE MFR IS NOT GUASCOR, CALC THE 25's 
                If SWavg = SWavg75 AndAlso SWavg75 = SWavg50 Then
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False) : SwCp75 = SwCp : SwCp50 = SwCp
                Else
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False)
                    SwCp75 = GetFluidValue(_SecCir_fluid.ToString, SWavg75, True, False)
                    SwCp50 = GetFluidValue(_SecCir_fluid.ToString, SWavg50, True, False)
                End If
            End If
        Else ' IF THE FLUID TYPE IS NOT WATER THEN GRAB THE PERCENT
            If _MFR = "Guascor" Then
                If SWavg = Swavg80 AndAlso Swavg80 = Swavg60 AndAlso Swavg60 = SWavg40 Then
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct) : SwCp80 = SwCp : SwCp60 = SwCp : SwCp40 = SwCp
                Else
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct)
                    SwCp80 = GetFluidValue(_SecCir_fluid.ToString, Swavg80, True, False, _f3pct)
                    SwCp60 = GetFluidValue(_SecCir_fluid.ToString, Swavg60, True, False, _f3pct)
                    SwCp40 = GetFluidValue(_SecCir_fluid.ToString, SWavg40, True, False, _f3pct)
                End If
            Else ' IF THE MFR IS NOT GUASCOR, CALC THE 25's 
                If SWavg = SWavg75 AndAlso SWavg75 = SWavg50 Then
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct) : SwCp75 = SwCp : SwCp50 = SwCp
                Else
                    SwCp = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct)
                    SwCp75 = GetFluidValue(_SecCir_fluid.ToString, SWavg75, True, False, _f3pct)
                    SwCp50 = GetFluidValue(_SecCir_fluid.ToString, SWavg50, True, False, _f3pct)
                End If
            End If
        End If

        ' FIND SW_DENSITY
        If _SecCir_fluid = FluidType.Water Then
            If _MFR = "Guascor" Then
                If SWavg = Swavg80 AndAlso Swavg80 = Swavg60 AndAlso Swavg60 = SWavg40 Then
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False) : SwDensity80 = SwDensity : SwDensity60 = SwDensity : SwDensity40 = SwDensity
                Else
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False)
                    SwDensity80 = GetFluidValue(_SecCir_fluid.ToString, Swavg80, True, False)
                    SwDensity60 = GetFluidValue(_SecCir_fluid.ToString, Swavg60, True, False)
                    SwDensity40 = GetFluidValue(_SecCir_fluid.ToString, SWavg40, True, False)
                End If
            Else ' IF THE MFR IS NOT GUASCOR, CALC THE 25's 
                If SWavg = SWavg75 AndAlso SWavg75 = SWavg50 Then
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False) : SwDensity75 = SwDensity : SwDensity50 = SwDensity
                Else
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False)
                    SwDensity75 = GetFluidValue(_SecCir_fluid.ToString, SWavg75, True, False)
                    SwDensity50 = GetFluidValue(_SecCir_fluid.ToString, SWavg50, True, False)
                End If
            End If
        Else ' IF THE FLUID TYPE IS NOT WATER THEN GRAB THE PERCENT
            If _MFR = "Guascor" Then
                If SWavg = Swavg80 AndAlso Swavg80 = Swavg60 AndAlso Swavg60 = SWavg40 Then
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct) : SwDensity80 = SwDensity : SwDensity60 = SwDensity : SwDensity40 = SwDensity
                Else
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct)
                    SwDensity80 = GetFluidValue(_SecCir_fluid.ToString, Swavg80, True, False, _f3pct)
                    SwDensity60 = GetFluidValue(_SecCir_fluid.ToString, Swavg60, True, False, _f3pct)
                    SwDensity40 = GetFluidValue(_SecCir_fluid.ToString, SWavg40, True, False, _f3pct)
                End If
            Else ' IF THE MFR IS NOT GUASCOR, CALC THE 25's 
                If SWavg = SWavg75 AndAlso SWavg75 = SWavg50 Then
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct) : SwDensity75 = SwDensity : SwDensity50 = SwDensity
                Else
                    SwDensity = GetFluidValue(_SecCir_fluid.ToString, SWavg, True, False, _f3pct)
                    SwDensity75 = GetFluidValue(_SecCir_fluid.ToString, SWavg75, True, False, _f3pct)
                    SwDensity50 = GetFluidValue(_SecCir_fluid.ToString, SWavg50, True, False, _f3pct)
                End If
            End If
        End If
    End Sub
#End Region
#Region "--> Uber Loop"
    Public Function UberLoop(solveFor As String, knownTemp As Double, fluid As FluidType, knownQorTemp As Double, knownGPM As Double) As Double
        ' T_in unknown && T_out known then subtract || T_in known && T_out unknown then add
        Dim counter As Integer = 0
        Dim unknown As Double = 0
        Dim testTemp As Double = 0
        Dim cp As Double = 0
        Dim density As Double = 0
        Dim percent As Integer = 0
        If Not fluid = FluidType.Water Then If fluid = _EngCoolant_fluid Then percent = _f1pct : If fluid = _PrmCir_fluid Then percent = _f2pct : If fluid = _SecCir_fluid Then percent = _f3pct

        Select Case solveFor
            Case "inlet"
                testTemp = knownTemp - knownQorTemp / (knownGPM * 500 * 0.85)
                avg = (testTemp + knownTemp) \ 2
                If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                unknown = knownTemp - knownQorTemp / (knownGPM * cp)
                ' ENTER LOOP
                While (counter < 5)
                    If unknown = testTemp Then Exit While
                    testTemp = unknown
                    avg = (testTemp + knownTemp) \ 2
                    If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                    unknown = knownTemp - knownQorTemp / (knownGPM * cp)
                    counter += 1
                End While
                Return unknown

            Case "outlet"
                testTemp = knownTemp + knownQorTemp / (knownGPM * 500 * 0.85)
                avg = (testTemp + knownTemp) \ 2
                If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                unknown = knownTemp + knownQorTemp / (knownGPM * cp)
                ' ENTER LOOP
                While (counter < 5)
                    If unknown = testTemp Then Exit While
                    testTemp = unknown
                    avg = (testTemp + knownTemp) \ 2
                    If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                    unknown = knownTemp + knownQorTemp / (knownGPM * cp)
                    counter += 1
                End While
                Return unknown

            Case "Q" ' not a loop calculation, but we use the same parameters to solve for it
                avg = (knownTemp + knownQorTemp) \ 2
                If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                If knownTemp > knownQorTemp Then
                    unknown = knownGPM * cp * (knownTemp - knownQorTemp)
                Else
                    unknown = knownGPM * cp * (knownQorTemp - knownTemp)
                End If
                Return unknown

            Case "guascorInlet"
                testTemp = knownTemp - knownQorTemp / (knownGPM * 500 * 0.85)
                avg = (testTemp + knownTemp) \ 2
                If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                If fluid = FluidType.Water Then density = GetFluidValue(fluid.ToString, avg, False, True) Else density = GetFluidValue(fluid.ToString, avg, False, True, percent)
                unknown = knownTemp - knownQorTemp / (knownGPM * cp * density * ConversionRatio)
                ' ENTER LOOP
                While (counter < 5)
                    If unknown = testTemp Then Exit While
                    testTemp = unknown
                    avg = (testTemp + knownTemp) \ 2
                    If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                    If fluid = FluidType.Water Then density = GetFluidValue(fluid.ToString, avg, False, True) Else density = GetFluidValue(fluid.ToString, avg, False, True, percent)
                    unknown = knownTemp - knownQorTemp / (knownGPM * cp * density * ConversionRatio)
                    counter += 1
                End While
                Return unknown

            Case "guascorOutlet"
                testTemp = knownTemp + knownQorTemp / (knownGPM * 500 * 0.85)
                avg = (testTemp + knownTemp) \ 2
                If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                If fluid = FluidType.Water Then density = GetFluidValue(fluid.ToString, avg, False, True) Else density = GetFluidValue(fluid.ToString, avg, False, True, percent)
                unknown = knownTemp + knownQorTemp / (knownGPM * cp * density * ConversionRatio)
                ' ENTER LOOP
                While (counter < 5)
                    If unknown = testTemp Then Exit While
                    testTemp = unknown
                    avg = (testTemp + knownTemp) \ 2
                    If fluid = FluidType.Water Then cp = GetFluidValue(fluid.ToString, avg, True, False) Else cp = GetFluidValue(fluid.ToString, avg, True, False, percent)
                    If fluid = FluidType.Water Then density = GetFluidValue(fluid.ToString, avg, False, True) Else density = GetFluidValue(fluid.ToString, avg, False, True, percent)
                    unknown = knownTemp + knownQorTemp / (knownGPM * cp * density * ConversionRatio)
                    counter += 1
                End While
                Return unknown
        End Select
        Return Nothing ' <-- this should never be called, this is just to satisfy all paths of this function
    End Function
#End Region

#Region "--> GPM's"
    Public Sub GetGPMs() ' THESE WILL REMAIN CONSTANT, NO PARTIALS!!
        avg = ((jw_in + jw_out) / 2)
        ' CP
        If _EngCoolant_fluid = FluidType.Water Then JWCp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False) Else JWCp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False, _f1pct)
        ' DENSITY
        If _EngCoolant_fluid = FluidType.Water Then JWdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True) Else JWdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True, _f1pct)
        ' IC SECTION
        avg = ((ic_in + ic_out) / 2)
        ' CP
        If _EngCoolant_fluid = FluidType.Water Then ICcp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False) Else ICcp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False, _f1pct)
        ' DENSITY
        If _EngCoolant_fluid = FluidType.Water Then ICdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True) Else ICdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True, _f1pct)

        ' CALC JWMassFlow
        If JWCp * (jw_out - jw_in) = 0 Then
            JWMassFlow = 1 : If _MFR = "MTU" Then jw_flow = 1
        Else
            If _MFR = "Guascor" Then
                JWMassFlow = jw_flow * ConversionRatio * JWdensity
            Else
                JWMassFlow = mainheat100 / (JWCp * (jw_out - jw_in)) : If _MFR = "MTU" Then jw_flow = JWMassFlow / (ConversionRatio * JWdensity)
            End If
        End If
        ' CALC ICMassFlow
        If ICcp * (ic_out - ic_in) = 0 Or ic_out = 1 Or ic_in = 1 Then
            ICMassFlow = 1 : If _MFR = "MTU" Then ic_flow = 1
        Else
            If _MFR = "Guascor" Then
                ICMassFlow = ic_flow * ConversionRatio * ICdensity
            Else
                ICMassFlow = lt_heat100 / (ICcp * (ic_out - ic_in)) : If _MFR = "MTU" Then ic_flow = ICMassFlow / (ConversionRatio * ICdensity)
            End If
        End If
    End Sub
#End Region
#Region "--> Steam"
    Public Sub CalcSteam()
        ' STEAM RECOVERED
        SQL.AddParam("@psig", _user_StmPress)
        SQL.ExecQuery("SELECT TOP(1) * FROM WaterPressure WHERE pressure>=@psig")
        If String.IsNullOrEmpty(SQL.Exception) Then
            SteamTemp = _get(SQL.DBDS, "temp") : VaporEnth = _get(SQL.DBDS, "vapor_enth")
        End If
        If ((SteamTemp + 100) >= _user_MinExTemp) Then
            QSteam = (exflow100 * ExCp) * (extemp100 - SteamTemp - PINCH_TEMP)
            If _MFR = "Guascor" Then
                QSteam80 = (exflow80 * ExCp) * (extemp80 - SteamTemp - PINCH_TEMP) : QSteam60 = (exflow60 * ExCp) * (extemp60 - SteamTemp - PINCH_TEMP) : QSteam40 = (exflow40 * ExCp) * (extemp40 - SteamTemp - PINCH_TEMP)
            Else
                QSteam75 = (exflow75 * ExCp) * (extemp75 - SteamTemp - PINCH_TEMP) : QSteam50 = (exflow50 * ExCp) * (extemp50 - SteamTemp - PINCH_TEMP)
            End If
        Else
            QSteam = ((exflow100 * ExCp) * (extemp100 - _user_MinExTemp - PINCH_TEMP))
            If _MFR = "Guascor" Then
                QSteam80 = (exflow80 * ExCp) * (extemp80 - _user_MinExTemp - PINCH_TEMP) : QSteam60 = (exflow60 * ExCp) * (extemp60 - _user_MinExTemp - PINCH_TEMP) : QSteam40 = (exflow40 * ExCp) * (extemp40 - _user_MinExTemp - PINCH_TEMP)
            Else
                QSteam75 = (exflow75 * ExCp) * (extemp75 - _user_MinExTemp - PINCH_TEMP) : QSteam50 = (exflow50 * ExCp) * (extemp50 - _user_MinExTemp - PINCH_TEMP)
            End If
        End If
        ' STEAM PRODUCTION
        SQL.AddParam("@feed", _user_Feed_H2O)
        SQL.ExecQuery("SELECT TOP(1) * FROM WaterTemp WHERE Temp>=@feed")
        If String.IsNullOrEmpty(SQL.Exception) Then
            SatLiq = _get(SQL.DBDS, "sat_liq")
        End If
        SteamProduction = QSteam / (VaporEnth - SatLiq)
        If _MFR = "Guascor" Then
            SteamProd80 = QSteam80 / (VaporEnth - SatLiq) : SteamProd60 = QSteam60 / (VaporEnth - SatLiq) : SteamProd40 = QSteam40 / (VaporEnth - SatLiq)
        Else
            SteamProd75 = QSteam75 / (VaporEnth - SatLiq) : SteamProd50 = QSteam50 / (VaporEnth - SatLiq)
        End If
    End Sub
#End Region
#Region "--> Efficiency/Fuelcon Calcs"
    Public Sub EfficiencyAndFuelConCalcs()
        ' FUEL CONSUMPTIONS
        btuKWh = fuelcon100 \ KWeOut100
        bHPhr = fuelcon100 \ engpow100
        If _MFR = "Guascor" Then
            btuKWh80 = fuelcon80 \ KWeOut80 : btuKWh60 = fuelcon60 \ KWeOut60 : btuKWh40 = fuelcon40 \ KWeOut40
            bHPhr80 = fuelcon80 \ engpow80 : bHPhr60 = fuelcon60 \ engpow60 : bHPhr40 = fuelcon40 \ engpow60
        Else
            btuKWh75 = fuelcon75 \ KWeOut75 : btuKW50 = fuelcon75 \ KWeOut50
            bHPhr75 = fuelcon75 \ engpow75 : bHPhr50 = fuelcon50 \ engpow50
        End If
        ' ELECTRICAL EFFICIENCY RATES
        EleEff = (KWeOut100 * 3412.1 / fuelcon100) * 100
        If _MFR = "Guascor" Then
            EleEff80 = (KWeOut80 * 3412.1 / fuelcon80) * 100 : EleEff60 = (KWeOut60 * 3412.1 / fuelcon60) * 100 : EleEff40 = (KWeOut40 * 3412.1 / fuelcon40) * 100
        Else
            If Not fuelcon75 = 0 Or Not fuelcon50 = 0 Then EleEff75 = (KWeOut75 * 3412.1 / fuelcon75) * 100 : EleEff50 = (KWeOut50 * 3412.1 / fuelcon50) * 100 Else EleEff75 = 0 : EleEff50 = 0
        End If
        ' THERMAL EFFICIENCY
        ThermEff = ((QHX + QICHX + QSteam) / fuelcon100) * 100
        If _MFR = "Guascor" Then
            ThermEff80 = ((QHX80 + QICHX80 + QSteam80) / fuelcon80) * 100 : ThermEff60 = ((QHX60 + QICHX60 + QSteam60) / fuelcon60) * 100 : ThermEff40 = ((QHX40 + QICHX40 + QSteam40) / fuelcon40) * 100
        Else
            ThermEff75 = ((QHX75 + QICHX75 + QSteam75) / fuelcon75) * 100 : ThermEff50 = ((QHX50 + QICHX50 + QSteam50) / fuelcon50) * 100
        End If
        ' TOTAL EFFICIENCY
        TotalEff = EleEff + ThermEff
        If _MFR = "Guascor" Then
            TotalEff80 = EleEff80 + ThermEff80 : TotalEff60 = EleEff60 + ThermEff60 : TotalEff40 = EleEff40 + ThermEff40
        Else
            TotalEff75 = EleEff75 + ThermEff75 : TotalEff50 = EleEff50 + ThermEff50
        End If
    End Sub
#End Region
#End Region
End Class