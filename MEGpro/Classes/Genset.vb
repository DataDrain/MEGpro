Public Class Genset
    Private SQL As New SQLControl : Private _query As String = ""

#Region "DECLARATIONS"
#Region "  engine stats"
    ' STANDARD STATS
    Public avg As Integer ' <-- this is an integer to prevent query errors when looking for decimals
    Public JWin As Integer : Public JWout As Integer : Public JWFlowRate As Double
    Public ICin As Integer : Public ICout As Integer : Public ICFlowRate As Double
    Public _ExFlow100 As Integer : Public _ExFlow80 As Integer : Public _ExFlow75 As Integer : Public _ExFlow60 As Integer : Public _ExFlow50 As Integer : Public _ExFlow40 As Integer
    Public _ExTemp100 As Integer : Public _ExTemp80 As Integer : Public _ExTemp75 As Integer : Public _ExTemp60 As Integer : Public _ExTemp50 As Integer : Public _ExTemp40 As Integer
    Public _HeatMain100u As Integer : Public _HeatMain80u As Integer : Public _HeatMain75u As Integer : Public _HeatMain60u As Integer : Public _HeatMain50u As Integer : Public _HeatMain40u As Integer
    Public _LTheat100u As Integer : Public _LTheat80u As Integer : Public _LTheat75u As Integer : Public _LTheat60u As Integer : Public _LTheat50u As Integer : Public _LTheat40u As Integer
    ' Note: FuelCon is aka EnergyIn
    Public _FuelCon100u As Integer : Public _FuelCon80u As Integer : Public _FuelCon75u As Integer : Public _FuelCon60u As Integer : Public _FuelCon50u As Integer : Public _FuelCon40u As Integer
    Public _OilCool100u As Integer : Public _OilCool80u As Integer : Public _OilCool60u As Integer : Public _OilCool40u As Integer
    Public elepow100 As Double : Public EngKW80 As Double : Public EngKW75 As Double : Public EngKW60 As Double : Public EngKW50 As Double : Public EngKW40 As Double
    Public btuKWh As Integer : Public btuKWh80 As Integer : Public btuKWh75 As Integer : Public btuKWh60 As Integer : Public btuKW50 As Integer : Public btuKWh40 As Integer
    Public bHPhr As Integer : Public bHPhr80 As Integer : Public bHPhr75 As Integer : Public bHPhr60 As Integer : Public bHPhr50 As Integer : Public bHPhr40 As Integer
    Public EngPow As Integer : Public EngPow80 As Integer : Public EngPow75 As Integer : Public EngPow60 As Integer : Public EngPow50 As Integer : Public EngPow40 As Integer
    Public EleEff As Double : Public EleEff80 As Double : Public EleEff75 As Double : Public EleEff60 As Double : Public EleEff50 As Double : Public EleEff40 As Double
    Public ThermEff As Double : Public ThermEff80 As Double : Public ThermEff75 As Double : Public ThermEff60 As Double : Public ThermEff50 As Double : Public ThermEff40 As Double
    Public TotalEff As Double : Public TotalEff80 As Double : Public TotalEff75 As Double : Public TotalEff60 As Double : Public TotalEff50 As Double : Public TotalEff40 As Double
    ' MAN STATS
    Public _CoolHeat100u As Integer : Public _CoolHeat75u As Integer : Public _CoolHeat50u As Integer
    Public _MixHT100u As Integer : Public _MixHT75u As Integer : Public _MixHT50u As Integer
    Public _Radiation100u As Integer : Public _Radiation75u As Integer : Public _Radiation50u As Integer
#End Region
#Region "  generator stats"
    Public strGenMFR As String
    Public _genID As String : Public _genRPM As Integer : Public _genKW As Double : Public _genKVA As Double : Public _genVolts As Integer
    Public x6 As Double : Public x5 As Double : Public x4 As Double
    Public x3 As Double : Public x2 As Double : Public x1 As Double : Public x0 As Double

    ' NEW VARS TO FIND TRUE GEN EFFICIENCY
    Public genEff As Double
    Public Const GEN_CONVERSION As Double = 0.7457
    Public genLoad As Double
    Public KWeOut As Double : Public KWeOut100 As Double : Public KWeOut80 As Double : Public KWeOut75 As Double : Public KWeOut60 As Double : Public KWeOut50 As Double : Public KWeOut40 As Double
    Public loopCount As Integer = 0
#End Region
#Region "  calculation vars"
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

    ' CONSTANTS
    Public Const PINCH_TEMP As Integer = 100
    Public Const ExCp As Double = 0.265 ' EX SPECIFIC HEAT
    Public Const CpBTU As Double = 0.85
    Public Const ConversionRatio As Double = 8.021
    Public Const MinApproachTemp As Double = 5

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
#Region "  from constructor"
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
    Public Sub New(eID As String, mfr As String, model As String, rpm As Integer, fuel As String, elepow As Double, pf As Single, _
                   minEx As Integer, stmPressure As Integer, feed As Integer, _
                   ppwIn As Integer, ppwOut As Integer, spwIn As Integer, spwOut As Integer, _
                   steam As Boolean, ehru As Boolean, eh2jw As Boolean, eh2primary As Boolean, _
                   jw As Boolean, lt As Boolean, lt2primary As Boolean, lt2second As Boolean, _
                   f1 As Integer, f2 As Integer, f3 As Integer, _
                   f1per As Integer, f2per As Integer, f3per As Integer, oil2jw As Boolean, oil2ic As Boolean)
        ' BASICS
        _EngID = eID : _MFR = mfr : _Model = model : _RPM = rpm : _Fuel = fuel : _PF = pf
        ' USER INPUTS
        _user_MinExTemp = minEx : _user_StmPress = stmPressure : _user_Feed_H2O = feed : _user_PWin = ppwIn : _user_PWout = ppwOut : _user_SWin = spwIn : _user_SWout = spwOut
        ' BOOLEANS
        _wantStm = steam : _wantEHRU = ehru : _EHRUtoJW = eh2jw : _EHRUtoPRM = eh2primary : _wantJW = jw : _wantLT = lt : _LTtoPRM = lt2primary : _LTtoSEC = lt2second
        If _MFR = "Guascor" Then _OilToJW = oil2jw : _OilToIC = oil2ic
        ' FLUID TYPES
        _EngCoolant_fluid = f1 : _f1pct = f1per : _PrmCir_fluid = f2 : _f2pct = f2per : _SecCir_fluid = f3 : _f3pct = f3per
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ' FIND CALC CASE BASED ON BOOLEANS
        FindCalcCase()

        'GetEngineStats()

        GetGeneratorStats()

        ' GET ENGINE INFO
        'Select Case _MFR
        '    Case "Guascor"
        '        GetGuascorData()
        '        If _OilToJW Then
        '            JWin = UberLoop("guascorInlet", JWout, _EngCoolant_fluid, (_HeatMain100u + _OilCool100u), JWFlowRate)
        '            ICout = UberLoop("guascorOutlet", ICin, _EngCoolant_fluid, _LTheat100u, ICFlowRate)
        '        End If
        '        If _OilToIC Then
        '            JWin = UberLoop("guascorInlet", JWout, _EngCoolant_fluid, _HeatMain100u, JWFlowRate)
        '            ICout = UberLoop("guascorOutlet", ICin, _EngCoolant_fluid, _LTheat100u + _OilCool100u, ICFlowRate)
        '        End If
        '    Case "MAN"
        '        'GetMANdata()
        '    Case "MTU"
        '        GetMTUdata()
        'End Select

        'GetGPMs()

        ' GET GENERATOR
        'If _GenID <> "" Then GetGenData() Else _GenID = "Included"
        'If _MFR <> "MTU" Then
        '    KWeOut100 = FindGenEfficiency(EnginePow)
        '    If _MFR = "Guascor" Then
        '        KWeOut80 = FindGenEfficiency(EnginePow80) : KWeOut60 = FindGenEfficiency(EnginePow60) : KWeOut40 = FindGenEfficiency(EnginePow40)
        '    Else
        '        KWeOut75 = FindGenEfficiency(EnginePow75) : KWeOut50 = FindGenEfficiency(EnginePow50)
        '    End If
        'End If

        '' MANDITORY CALCS
        'QExAvail = _ExFlow100 * ExCp * (_ExTemp100 - _user_MinExTemp)
        'If _MFR = "Guascor" Then
        '    QExAvail80 = _ExFlow80 * ExCp * (_ExTemp80 - _user_MinExTemp) : QExAvail60 = _ExFlow60 * ExCp * (_ExTemp60 - _user_MinExTemp) : QExAvail40 = _ExFlow40 * ExCp * (_ExTemp40 - _user_MinExTemp)
        'Else
        '    QExAvail75 = _ExFlow75 * ExCp * (_ExTemp75 - _user_MinExTemp) : QExAvail50 = _ExFlow50 * ExCp * (_ExTemp50 - _user_MinExTemp)
        'End If

        '' CONDITIONAL CALCS
        'If _wantStm = True Then CalcSteam()
        'If _wantEHRU = True Then
        '    QEHRU = QExAvail - QSteam
        '    If _MFR = "Guascor" Then
        '        QEHRU80 = QExAvail80 - QSteam80 : QEHRU60 = QExAvail60 - QSteam60 : QEHRU40 = QExAvail40 - QSteam40
        '    Else
        '        QEHRU75 = QExAvail75 - QSteam75 : QEHRU50 = QExAvail50 - QSteam50
        '    End If
        'End If

        'CaseCalculations()

        'EfficiencyAndFuelConCalcs()
    End Sub
#End Region

#Region "QUERY SUBS"
#Region "--> Engine Stats"
    Public Sub GetEngineStats()
        Select Case _MFR
            Case "Guascor"
                _query = String.Format("SELECT {0} FROM Engines", GUASCOR_GENSET)
                'EngKW80 = _get(SQL.DBDS, "elepow80")
                MsgBox(_query)
            Case "MTU"
        End Select
    End Sub
#End Region
#Region "  guascor"
    Public Sub GetGuascorData() ' GUASCOR
        SQL.AddParam("@eid", _EngID)
        _query = String.Format("SELECT * FROM {0} WHERE Engine_ID=@eid", _MFR)
        SQL.ExecQuery(_query)
        If String.IsNullOrEmpty(SQL.Exception) Then
            _Model = SQL.DBDS.Tables(0).Rows(0)("Model")
            _Fuel = SQL.DBDS.Tables(0).Rows(0)("FuelType") : _BurnType = SQL.DBDS.Tables(0).Rows(0)("BurnType")
            _RPM = SQL.DBDS.Tables(0).Rows(0)("RPM")
            elepow100 = SQL.DBDS.Tables(0).Rows(0)("ElePow100") : EngKW80 = SQL.DBDS.Tables(0).Rows(0)("ElePow80") : EngKW60 = SQL.DBDS.Tables(0).Rows(0)("ElePow60") : EngKW40 = SQL.DBDS.Tables(0).Rows(0)("ElePow40")
            EngPow = SQL.DBDS.Tables(0).Rows(0)("MechPow100") : EngPow80 = SQL.DBDS.Tables(0).Rows(0)("MechPow80") : EngPow60 = SQL.DBDS.Tables(0).Rows(0)("MechPow60") : EngPow40 = SQL.DBDS.Tables(0).Rows(0)("MechPow40")
            JWout = SQL.DBDS.Tables(0).Rows(0)("JWTemp")
            ICin = SQL.DBDS.Tables(0).Rows(0)("ICTemp")
            JWFlowRate = SQL.DBDS.Tables(0).Rows(0)("JWFlow")
            ICFlowRate = SQL.DBDS.Tables(0).Rows(0)("ICFlow")
            _ExFlow100 = SQL.DBDS.Tables(0).Rows(0)("ExFlow100") : _ExFlow80 = SQL.DBDS.Tables(0).Rows(0)("ExFlow80") : _ExFlow60 = SQL.DBDS.Tables(0).Rows(0)("ExFlow60") : _ExFlow40 = SQL.DBDS.Tables(0).Rows(0)("ExFlow40")
            _ExTemp100 = SQL.DBDS.Tables(0).Rows(0)("ExTemp100") : _ExTemp80 = SQL.DBDS.Tables(0).Rows(0)("ExTemp80") : _ExTemp60 = SQL.DBDS.Tables(0).Rows(0)("ExTemp60") : _ExTemp40 = SQL.DBDS.Tables(0).Rows(0)("ExTemp40")
            _FuelCon100u = SQL.DBDS.Tables(0).Rows(0)("FuelCon100u") : _FuelCon80u = SQL.DBDS.Tables(0).Rows(0)("FuelCon80u") : _FuelCon60u = SQL.DBDS.Tables(0).Rows(0)("FuelCon60u") : _FuelCon40u = SQL.DBDS.Tables(0).Rows(0)("FuelCon40u")
            _HeatMain100u = SQL.DBDS.Tables(0).Rows(0)("HeatMain100u") : _HeatMain80u = SQL.DBDS.Tables(0).Rows(0)("HeatMain80u") : _HeatMain60u = SQL.DBDS.Tables(0).Rows(0)("HeatMain60u") : _HeatMain40u = SQL.DBDS.Tables(0).Rows(0)("HeatMain40u")
            _LTheat100u = SQL.DBDS.Tables(0).Rows(0)("LTheat100u") : _LTheat80u = SQL.DBDS.Tables(0).Rows(0)("LTheat80u") : _LTheat60u = SQL.DBDS.Tables(0).Rows(0)("LTheat60u") : _LTheat40u = SQL.DBDS.Tables(0).Rows(0)("LTheat40u")
            _OilCool100u = SQL.DBDS.Tables(0).Rows(0)("OilCool100u") : _OilCool80u = SQL.DBDS.Tables(0).Rows(0)("OilCool80u") : _OilCool60u = SQL.DBDS.Tables(0).Rows(0)("OilCool60u") : _OilCool40u = SQL.DBDS.Tables(0).Rows(0)("OilCool40u")
        Else
            MsgBox(SQL.Exception)
        End If
    End Sub
#End Region
#Region "  mtu"
    Public Sub GetMTUdata() ' MTU
        SQL.AddParam("@eid", _EngID)
        _query = String.Format("SELECT * FROM {0} WHERE Engine_ID=@eid", _MFR)
        SQL.ExecQuery(_query)
        If String.IsNullOrEmpty(SQL.Exception) Then
            _Model = SQL.DBDS.Tables(0).Rows(0)("Model")
            _Fuel = SQL.DBDS.Tables(0).Rows(0)("FuelType") : _BurnType = SQL.DBDS.Tables(0).Rows(0)("BurnType")
            _RPM = SQL.DBDS.Tables(0).Rows(0)("RPM")
            _genVolts = SQL.DBDS.Tables(0).Rows(0)("Voltage")
            _genKW = SQL.DBDS.Tables(0).Rows(0)("ElePow100")
            KWeOut100 = SQL.DBDS.Tables(0).Rows(0)("ElePow100") : KWeOut75 = SQL.DBDS.Tables(0).Rows(0)("ElePow75") : KWeOut50 = SQL.DBDS.Tables(0).Rows(0)("ElePow50")
            elepow100 = _genKW : EngKW75 = KWeOut75 : EngKW50 = KWeOut50
            EngPow = SQL.DBDS.Tables(0).Rows(0)("EnginePow100") : EngPow75 = SQL.DBDS.Tables(0).Rows(0)("EnginePow75") : EngPow50 = SQL.DBDS.Tables(0).Rows(0)("EnginePow50")
            JWin = SQL.DBDS.Tables(0).Rows(0)("JWin")
            JWout = SQL.DBDS.Tables(0).Rows(0)("JWout")
            ICin = SQL.DBDS.Tables(0).Rows(0)("ICin")
            ICout = SQL.DBDS.Tables(0).Rows(0)("ICout")
            _ExFlow100 = SQL.DBDS.Tables(0).Rows(0)("ExFlow100") : _ExFlow75 = SQL.DBDS.Tables(0).Rows(0)("ExFlow75") : _ExFlow50 = SQL.DBDS.Tables(0).Rows(0)("ExFlow50")
            _ExTemp100 = SQL.DBDS.Tables(0).Rows(0)("ExTemp100") : _ExTemp75 = SQL.DBDS.Tables(0).Rows(0)("ExTemp75") : _ExTemp50 = SQL.DBDS.Tables(0).Rows(0)("ExTemp50")
            _FuelCon100u = SQL.DBDS.Tables(0).Rows(0)("EnergyIn100u") : _FuelCon75u = SQL.DBDS.Tables(0).Rows(0)("EnergyIn75u") : _FuelCon50u = SQL.DBDS.Tables(0).Rows(0)("EnergyIn50u")
            _HeatMain100u = SQL.DBDS.Tables(0).Rows(0)("HeatMain100u") : _HeatMain75u = SQL.DBDS.Tables(0).Rows(0)("HeatMain75u") : _HeatMain50u = SQL.DBDS.Tables(0).Rows(0)("HeatMain50u")
            _LTheat100u = SQL.DBDS.Tables(0).Rows(0)("LTheat100u") : _LTheat75u = SQL.DBDS.Tables(0).Rows(0)("LTheat75u") : _LTheat50u = SQL.DBDS.Tables(0).Rows(0)("LTheat50u")
        Else
            MsgBox(SQL.Exception)
        End If
    End Sub
#End Region
    Public Sub GetGeneratorStats()
        SQL.AddParam("@id", _EngID)
        SQL.ExecQuery("SELECT TOP(1) g.id FROM generators AS g, engines AS e WHERE e.id=@id AND e.rpm = g.rpm AND g.elepow > e.elepow100")
        _GenID = _get(SQL.DBDS, "id")
        MsgBox(_GenID)
        'SQL.AddParam("@gen", _GenID)
        '_query = "SELECT * FROM Generators WHERE Gen_ID=@gen"
        'SQL.ExecQuery(_query)
        'If String.IsNullOrEmpty(SQL.Exception) Then
        '    _GenID = String.Format("Newage - {0}", SQL.DBDS.Tables(0).Rows(0)("Gen_ID"))
        '    'strGenMFR = SQL.DBDS.Tables(0).Rows(0)("MFR")
        '    _genKVA = SQL.DBDS.Tables(0).Rows(0)("kVA")
        '    _genKW = SQL.DBDS.Tables(0).Rows(0)("ElePow100")
        '    _genVolts = SQL.DBDS.Tables(0).Rows(0)("Voltage")
        '    Select Case _PF
        '        Case 0.8
        '            x6 = SQL.DBDS.Tables(0).Rows(0)("Pf8x6")
        '            x5 = SQL.DBDS.Tables(0).Rows(0)("Pf8x5")
        '            x4 = SQL.DBDS.Tables(0).Rows(0)("Pf8x4")
        '            x3 = SQL.DBDS.Tables(0).Rows(0)("Pf8x3")
        '            x2 = SQL.DBDS.Tables(0).Rows(0)("Pf8x2")
        '            x1 = SQL.DBDS.Tables(0).Rows(0)("Pf8x1")
        '            x0 = SQL.DBDS.Tables(0).Rows(0)("Pf8x0")
        '        Case 0.9
        '            x6 = SQL.DBDS.Tables(0).Rows(0)("Pf9x6")
        '            x5 = SQL.DBDS.Tables(0).Rows(0)("Pf9x5")
        '            x4 = SQL.DBDS.Tables(0).Rows(0)("Pf9x4")
        '            x3 = SQL.DBDS.Tables(0).Rows(0)("Pf9x3")
        '            x2 = SQL.DBDS.Tables(0).Rows(0)("Pf9x2")
        '            x1 = SQL.DBDS.Tables(0).Rows(0)("Pf9x1")
        '            x0 = SQL.DBDS.Tables(0).Rows(0)("Pf9x0")
        '        Case 1.0
        '            x6 = SQL.DBDS.Tables(0).Rows(0)("Pf1x6")
        '            x5 = SQL.DBDS.Tables(0).Rows(0)("Pf1x5")
        '            x4 = SQL.DBDS.Tables(0).Rows(0)("Pf1x4")
        '            x3 = SQL.DBDS.Tables(0).Rows(0)("Pf1x3")
        '            x2 = SQL.DBDS.Tables(0).Rows(0)("Pf1x2")
        '            x1 = SQL.DBDS.Tables(0).Rows(0)("Pf1x1")
        '            x0 = SQL.DBDS.Tables(0).Rows(0)("Pf1x0")
        '    End Select
        'Else
        '    MsgBox(SQL.Exception)
        'End If
    End Sub

#Region "  circuit fluids"
    Public Function GetFluidValue(type As String, Temp As Double, Cp As Boolean, Density As Boolean, Optional Percent As Double = Nothing) As Double
        Dim tblName As String = "" : Dim tblPrefix As String = "" : Dim tblSuffix As String = ""
        Dim colName As String = ""

        ' PREPARE THE QUERY BASED UPON PARAMETERS
        If type = "Ethylene" Or type = "Propylene" Then
            tblPrefix = type
            If Cp = True Then tblSuffix = "Cp"
            If Density = True Then tblSuffix = "Density"
            tblName = tblPrefix & tblSuffix
            colName = "P" & Percent
            _query = String.Format("SELECT {0} FROM {1} WHERE Temp>={2} ORDER BY TEMP", colName, tblName, Temp)
        ElseIf type = "Water" Then
            tblPrefix = type : tblSuffix = "Properties"
            tblName = tblPrefix & tblSuffix
            If Cp = True Then colName = "Cp"
            If Density = True Then colName = "Density"
        Else : Return 0 : End If

        _query = String.Format("SELECT {0} FROM {1} WHERE Temp>={2} ORDER BY TEMP", colName, tblName, Temp)
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
#Region "  calc cases"
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
#Region "  primary pw"
    Public Sub DeterminePWtemps()
        If _MFR = "Guascor" Then ' inlet will vary, JWout is constant
            If _OilToJW Then
                jwin80 = UberLoop("inlet", JWout, _EngCoolant_fluid, (_HeatMain80u + _OilCool80u), JWMassFlow) : jwin60 = UberLoop("inlet", JWout, _EngCoolant_fluid, (_HeatMain60u + _OilCool60u), JWMassFlow) : jwin40 = UberLoop("inlet", JWout, _EngCoolant_fluid, (_HeatMain40u + _OilCool40u), JWMassFlow)
            Else
                jwin80 = UberLoop("inlet", JWout, _EngCoolant_fluid, _HeatMain80u, JWMassFlow) : jwin60 = UberLoop("inlet", JWout, _EngCoolant_fluid, _HeatMain60u, JWMassFlow) : jwin40 = UberLoop("inlet", JWout, _EngCoolant_fluid, _HeatMain40u, JWMassFlow)
            End If
            ' DETERMINE IF PRIMARY PW TEMPS ARE ACCEPTABLE
            If Not JWMassFlow = 0 Or Not JWCp = 0 Then
                PostEHRU = UberLoop("outlet", JWout, _EngCoolant_fluid, QEHRU, JWMassFlow)
                PostEHRU80 = UberLoop("outlet", JWout, _EngCoolant_fluid, QEHRU80, JWMassFlow) : PostEHRU60 = UberLoop("outlet", JWout, _EngCoolant_fluid, QEHRU60, JWMassFlow) : PostEHRU40 = UberLoop("outlet", JWout, _EngCoolant_fluid, QEHRU40, JWMassFlow)
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
            If JWin - _user_PWin < MinApproachTemp Then PostHX = PwInActual + MinApproachTemp Else PostHX = JWin
            If jwin80 - _user_PWin < MinApproachTemp Then PostHX80 = PwInActual80 + MinApproachTemp Else PostHX80 = jwin80
            If jwin60 - _user_PWin < MinApproachTemp Then PostHX60 = PwInActual60 + MinApproachTemp Else PostHX60 = jwin60
            If jwin40 - _user_PWin < MinApproachTemp Then PostHX40 = PwInActual40 + MinApproachTemp Else PostHX40 = jwin40
        Else '==== 25's ====
            jwout75 = UberLoop("outlet", JWin, _EngCoolant_fluid, _HeatMain75u, JWMassFlow) : jwout50 = UberLoop("outlet", JWin, _EngCoolant_fluid, _HeatMain50u, JWMassFlow)
            ' DETERMINES IF PRIMARY PW TEMPS ARE ACCEPTABLE
            If Not JWMassFlow = 0 Or Not JWCp = 0 Then
                PostEHRU = UberLoop("outlet", JWout, _EngCoolant_fluid, QEHRU, JWMassFlow) ' <-- PostEHRU = QEHRU / (JwMassFlow * JwCp) + JwOut
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
            If JWin - _user_PWin < MinApproachTemp Then
                PostHX = PwInActual + MinApproachTemp : PostHX75 = PwInActual75 + MinApproachTemp : PostHX50 = PwInActual50 + MinApproachTemp
            Else
                PostHX = JWin : PostHX75 = JWin : PostHX50 = JWin
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
        QJWRad = UberLoop("Q", PostHX, _EngCoolant_fluid, JWin, JWMassFlow) ' <--- 'QJWRad = JWMassFlow * JWCp * (PostHX - JWin)
        If _MFR = "Guascor" Then
            QJWRad80 = UberLoop("Q", PostHX80, _EngCoolant_fluid, jwin80, JWMassFlow) : QJWRad60 = UberLoop("Q", PostHX60, _EngCoolant_fluid, jwin60, JWMassFlow) : QJWRad40 = UberLoop("Q", PostHX40, _EngCoolant_fluid, jwin40, JWMassFlow)
        Else
            QJWRad75 = UberLoop("Q", PostHX75, _EngCoolant_fluid, JWin, JWMassFlow) : QJWRad50 = UberLoop("Q", PostHX50, _EngCoolant_fluid, JWin, JWMassFlow)
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
#Region "  secondary pw"
    Public Sub DetermineSWtemps()
        If _MFR = "Guascor" Then
            icout80 = UberLoop("outlet", ICin, _SecCir_fluid, _LTheat80u, ICMassFlow) : icout60 = UberLoop("outlet", ICin, _SecCir_fluid, _LTheat60u, ICMassFlow) : icout40 = UberLoop("outlet", ICin, _SecCir_fluid, _LTheat40u, ICMassFlow)
        Else
            icout75 = UberLoop("outlet", ICin, _EngCoolant_fluid, _LTheat75u, ICMassFlow) : icout50 = UberLoop("outlet", ICin, _EngCoolant_fluid, _LTheat50u, ICMassFlow)
        End If
        ' SW_OUT ACTUALS
        If ICout - _user_SWout < MinApproachTemp Then SwOutActual = ICout - MinApproachTemp : tooHot = True Else SwOutActual = _user_SWout
        If _MFR = "Guascor" Then
            If icout80 - _user_SWout < MinApproachTemp Then SwOutActual80 = icout80 - MinApproachTemp : tooHot = True Else SwOutActual80 = _user_SWout
            If icout60 - _user_SWout < MinApproachTemp Then SwOutActual60 = icout60 - MinApproachTemp : tooHot = True Else SwOutActual60 = _user_SWout
            If icout40 - _user_SWout < MinApproachTemp Then SwOutActual40 = icout40 - MinApproachTemp : tooHot = True Else SwOutActual40 = _user_SWout
        Else
            If icout75 - _user_SWout < MinApproachTemp Then SwOutActual75 = icout75 - MinApproachTemp : tooHot = True Else SwOutActual75 = _user_SWout
            If icout50 - _user_SWout < MinApproachTemp Then SwOutActual50 = icout50 - MinApproachTemp : tooHot = True Else SwOutActual50 = _user_SWout
        End If
        ' SW_IN ACTUALS
        If ICout - _user_SWin < MinApproachTemp Then SwInActual = ICout - MinApproachTemp : tooHot = True Else SwInActual = _user_SWin
        If _MFR = "Guascor" Then
            If icout80 - _user_SWin < MinApproachTemp Then SwInActual80 = icout80 - MinApproachTemp : tooHot = True Else SwInActual80 = _user_SWin
            If icout60 - _user_SWin < MinApproachTemp Then SwInActual60 = icout60 - MinApproachTemp : tooHot = True Else SwInActual60 = _user_SWin
            If icout40 - _user_SWin < MinApproachTemp Then SwInActual40 = icout40 - MinApproachTemp : tooHot = True Else SwInActual40 = _user_SWin
        Else
            If icout75 - _user_SWin < MinApproachTemp Then SwInActual75 = icout75 - MinApproachTemp : tooHot = True Else SwInActual75 = _user_SWin
            If icout50 - _user_SWin < MinApproachTemp Then SwInActual50 = icout50 - MinApproachTemp : tooHot = True Else SwInActual50 = _user_SWin
        End If
        ' POST IC_HX
        If ICin - _user_SWin < MinApproachTemp Then
            PostICHX = SwInActual + MinApproachTemp
            If _MFR = "Guascor" Then
                PostICHX80 = SwInActual80 + MinApproachTemp : PostICHX60 = SwInActual60 + MinApproachTemp : PostICHX40 = SwInActual40 + MinApproachTemp
            Else
                PostICHX75 = SwInActual75 + MinApproachTemp : PostICHX50 = SwInActual50 + MinApproachTemp
            End If
            If PostICHX >= ICout Then PostICHX = ICout
            If _MFR = "Guascor" Then
                If PostICHX80 >= icout80 Then PostICHX80 = icout80 : If PostICHX60 >= icout80 Then PostICHX60 = icout60 : If PostICHX40 >= icout80 Then PostICHX40 = icout40
            Else
                If PostICHX75 >= icout75 Then PostICHX75 = icout75 : If PostICHX50 >= icout50 Then PostICHX50 = icout50
            End If
        Else
            PostICHX = ICin
            If _MFR = "Guascor" Then
                PostICHX80 = ICin : PostICHX60 = ICin : PostICHX40 = ICin
            Else
                PostICHX75 = ICin : PostICHX50 = ICin
            End If
        End If
    End Sub

    Public Sub CalcSWflow() ' CALC HEAT TRANSFERRED THROUGH HEAT EXCHANGER & HEAT LOST TO RADIATOR
        ' QICHX
        QICHX = UberLoop("Q", ICin, _SecCir_fluid, ICout, ICMassFlow)
        If _MFR = "Guascor" Then
            QICHX80 = UberLoop("Q", ICin, _SecCir_fluid, icout80, ICMassFlow) : QICHX60 = UberLoop("Q", ICin, _SecCir_fluid, icout60, ICMassFlow) : QICHX40 = UberLoop("Q", ICin, _SecCir_fluid, icout40, ICMassFlow)
        Else
            QICHX75 = UberLoop("Q", ICin, _SecCir_fluid, icout75, ICMassFlow) : QICHX50 = UberLoop("Q", ICin, _SecCir_fluid, icout50, ICMassFlow)
        End If
        'QICRAD
        QICRad = UberLoop("Q", PostICHX, _SecCir_fluid, ICin, ICMassFlow)
        If _MFR = "Guascor" Then
            QICRad80 = UberLoop("Q", PostICHX80, _SecCir_fluid, ICin, ICMassFlow) : QICRad60 = UberLoop("Q", PostICHX60, _SecCir_fluid, ICin, ICMassFlow) : QICRad40 = UberLoop("Q", PostICHX40, _SecCir_fluid, ICin, ICMassFlow)
        Else
            QICRad75 = UberLoop("Q", PostICHX75, _SecCir_fluid, ICin, ICMassFlow) : QICRad50 = UberLoop("Q", PostICHX50, _SecCir_fluid, ICin, ICMassFlow)
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
#Region "  uber loop"
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

            Case "Q"
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
#Region "  gen efficiency"
    Private Function FindGenEfficiency(bhp As Double) As Double
        Dim kwEtest As Double = bhp * GEN_CONVERSION * 0.95

        genLoad = kwEtest / (_PF * _genKVA)
        genEff = ((genLoad ^ 6 * x6) + (genLoad ^ 5 * x5) + (genLoad ^ 4 * x4) + (genLoad ^ 3 * x3) + (genLoad ^ 2 * x2) + (genLoad ^ 1 * x1) + (x0)) / 100
        KWeOut = bhp * GEN_CONVERSION * genEff

        ' ENTER LOOP
        While (loopCount < 5)
            If System.Math.Abs(kwEtest - KWeOut) <= 0.5 Then Exit While
            kwEtest = KWeOut
            genLoad = kwEtest / (_PF * _genKVA)
            genEff = ((genLoad ^ 6 * x6) + (genLoad ^ 5 * x5) + (genLoad ^ 4 * x4) + (genLoad ^ 3 * x3) + (genLoad ^ 2 * x2) + (genLoad ^ 1 * x1) + (x0)) / 100
            KWeOut = bhp * GEN_CONVERSION * genEff
            loopCount += 1
        End While
        Return KWeOut
    End Function
#End Region
#Region "  gpm's"
    Public Sub GetGPMs() ' THESE WILL REMAIN CONSTANT, NO PARTIALS!!
        avg = ((JWin + JWout) / 2)
        ' CP
        If _EngCoolant_fluid = FluidType.Water Then JWCp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False) Else JWCp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False, _f1pct)
        ' DENSITY
        If _EngCoolant_fluid = FluidType.Water Then JWdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True) Else JWdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True, _f1pct)
        ' IC SECTION
        avg = ((ICin + ICout) / 2)
        ' CP
        If _EngCoolant_fluid = FluidType.Water Then ICcp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False) Else ICcp = GetFluidValue(_EngCoolant_fluid.ToString, avg, True, False, _f1pct)
        ' DENSITY
        If _EngCoolant_fluid = FluidType.Water Then ICdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True) Else ICdensity = GetFluidValue(_EngCoolant_fluid.ToString, avg, False, True, _f1pct)

        ' CALC JWMassFlow
        If JWCp * (JWout - JWin) = 0 Then
            JWMassFlow = 1 : If _MFR = "MTU" Then JWFlowRate = 1
        Else
            If _MFR = "Guascor" Then
                JWMassFlow = JWFlowRate * ConversionRatio * JWdensity
            Else
                JWMassFlow = _HeatMain100u / (JWCp * (JWout - JWin)) : If _MFR = "MTU" Then JWFlowRate = JWMassFlow / (ConversionRatio * JWdensity)
            End If
        End If
        ' CALC ICMassFlow
        If ICcp * (ICout - ICin) = 0 Or ICout = 1 Or ICin = 1 Then
            ICMassFlow = 1 : If _MFR = "MTU" Then ICFlowRate = 1
        Else
            If _MFR = "Guascor" Then
                ICMassFlow = ICFlowRate * ConversionRatio * ICdensity
            Else
                ICMassFlow = _LTheat100u / (ICcp * (ICout - ICin)) : If _MFR = "MTU" Then ICFlowRate = ICMassFlow / (ConversionRatio * ICdensity)
            End If
        End If
    End Sub
#End Region
#Region "  steam"
    Public Sub CalcSteam()
        ' STEAM RECOVERED
        SQL.AddParam("@psig", _user_StmPress)
        _query = ("SELECT * FROM WaterPressure WHERE SteamPressure>=@psig ORDER BY SteamPressure")
        SQL.ExecQuery(_query)
        If String.IsNullOrEmpty(SQL.Exception) Then
            SteamTemp = SQL.DBDS.Tables(0).Rows(0)("SteamTemp") : VaporEnth = SQL.DBDS.Tables(0).Rows(0)("VaporEnth")
        End If
        If ((SteamTemp + 100) >= _user_MinExTemp) Then
            QSteam = (_ExFlow100 * ExCp) * (_ExTemp100 - SteamTemp - PINCH_TEMP)
            If _MFR = "Guascor" Then
                QSteam80 = (_ExFlow80 * ExCp) * (_ExTemp80 - SteamTemp - PINCH_TEMP) : QSteam60 = (_ExFlow60 * ExCp) * (_ExTemp60 - SteamTemp - PINCH_TEMP) : QSteam40 = (_ExFlow40 * ExCp) * (_ExTemp40 - SteamTemp - PINCH_TEMP)
            Else
                QSteam75 = (_ExFlow75 * ExCp) * (_ExTemp75 - SteamTemp - PINCH_TEMP) : QSteam50 = (_ExFlow50 * ExCp) * (_ExTemp50 - SteamTemp - PINCH_TEMP)
            End If
        Else
            QSteam = ((_ExFlow100 * ExCp) * (_ExTemp100 - _user_MinExTemp - PINCH_TEMP))
            If _MFR = "Guascor" Then
                QSteam80 = (_ExFlow80 * ExCp) * (_ExTemp80 - _user_MinExTemp - PINCH_TEMP) : QSteam60 = (_ExFlow60 * ExCp) * (_ExTemp60 - _user_MinExTemp - PINCH_TEMP) : QSteam40 = (_ExFlow40 * ExCp) * (_ExTemp40 - _user_MinExTemp - PINCH_TEMP)
            Else
                QSteam75 = (_ExFlow75 * ExCp) * (_ExTemp75 - _user_MinExTemp - PINCH_TEMP) : QSteam50 = (_ExFlow50 * ExCp) * (_ExTemp50 - _user_MinExTemp - PINCH_TEMP)
            End If
        End If
        ' STEAM PRODUCTION
        SQL.AddParam("@feed", _user_Feed_H2O)
        _query = ("SELECT * FROM WaterTemp WHERE Temp>=@feed ORDER BY Temp")
        SQL.ExecQuery(_query)
        If String.IsNullOrEmpty(SQL.Exception) Then
            SatLiq = SQL.DBDS.Tables(0).Rows(0)("SatLiq")
        End If
        SteamProduction = QSteam / (VaporEnth - SatLiq)
        If _MFR = "Guascor" Then
            SteamProd80 = QSteam80 / (VaporEnth - SatLiq) : SteamProd60 = QSteam60 / (VaporEnth - SatLiq) : SteamProd40 = QSteam40 / (VaporEnth - SatLiq)
        Else
            SteamProd75 = QSteam75 / (VaporEnth - SatLiq) : SteamProd50 = QSteam50 / (VaporEnth - SatLiq)
        End If
    End Sub
#End Region
#Region "  efficiency calcs"
    Public Sub EfficiencyAndFuelConCalcs()
        ' FUEL CONSUMPTIONS
        btuKWh = _FuelCon100u \ KWeOut100
        bHPhr = _FuelCon100u \ EngPow
        If _MFR = "Guascor" Then
            btuKWh80 = _FuelCon80u \ EngKW80 : btuKWh60 = _FuelCon60u \ EngKW60 : btuKWh40 = _FuelCon40u \ EngKW40
            bHPhr80 = _FuelCon80u \ EngPow80 : bHPhr60 = _FuelCon60u \ EngPow60 : bHPhr40 = _FuelCon40u \ EngPow60
        Else
            btuKWh75 = _FuelCon75u \ EngKW75 : btuKW50 = _FuelCon75u \ EngKW50
            bHPhr75 = _FuelCon75u \ EngPow75 : bHPhr50 = _FuelCon50u \ EngPow50
        End If
        ' ELECTRICAL EFFICIENCY RATES
        EleEff = (KWeOut100 * 3412.1 / _FuelCon100u) * 100
        If _MFR = "Guascor" Then
            EleEff80 = (KWeOut80 * 3412.1 / _FuelCon80u) * 100 : EleEff60 = (KWeOut60 * 3412.1 / _FuelCon60u) * 100 : EleEff40 = (KWeOut40 * 3412.1 / _FuelCon40u) * 100
        Else
            If Not _FuelCon75u = 0 Or Not _FuelCon50u = 0 Then EleEff75 = (KWeOut75 * 3412.1 / _FuelCon75u) * 100 : EleEff50 = (KWeOut50 * 3412.1 / _FuelCon50u) * 100 Else EleEff75 = 0 : EleEff50 = 0
        End If
        ' THERMAL EFFICIENCY
        ThermEff = ((QHX + QICHX + QSteam) / _FuelCon100u) * 100
        If _MFR = "Guascor" Then
            ThermEff80 = ((QHX80 + QICHX80 + QSteam80) / _FuelCon80u) * 100 : ThermEff60 = ((QHX60 + QICHX60 + QSteam60) / _FuelCon60u) * 100 : ThermEff40 = ((QHX40 + QICHX40 + QSteam40) / _FuelCon40u) * 100
        Else
            ThermEff75 = ((QHX75 + QICHX75 + QSteam75) / _FuelCon75u) * 100 : ThermEff50 = ((QHX50 + QICHX50 + QSteam50) / _FuelCon50u) * 100
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

    Public Sub GensetMsgBox()
        Dim FinalOutput As String = Nothing
        FinalOutput = String.Format("Eng ID = {1} // RPM = {2} // ElePow100 = {3}{0}" +
                                 "Gen ID = {4} // g_ElePow = {5} // g_RPM = {6}{0} CalcCase = {7}", Environment.NewLine, _
                                _EngID, _RPM, elepow100, _GenID, _genKW, _genRPM, CalcCase)
        MsgBox(FinalOutput)
    End Sub
End Class