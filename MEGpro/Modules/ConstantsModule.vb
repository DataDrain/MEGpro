Module ConstantsModule
#Region "STRING CONSTANTS"
    Private Const STANDARD_STATS As String = "id, mfr, model, rpm, fuel, burn_type, nox, comp_ratio, ignition_time, jw_flow, ic_flow"

    Public Const ALL_GUASCOR As String = "id, model, rpm, fuel, burn_type, nox, comp_ratio, ignition_time, jw_flow, ic_flow, dex, jw_out, ic_in, " & _
                                      "elepow100, elepow80, elepow60, elepow40, engpow100, engpow80, engpow60, engpow40, mainheat100, mainheat80, mainheat60, mainheat40, " & _
                                      "lt_heat100, lt_heat80, lt_heat60, lt_heat40, exflow100, exflow80, exflow60, exflow40, extemp100, extemp80, extemp60, extemp40, " & _
                                      "fuelcon100, fuelcon80, fuelcon60, fuelcon40, heat_radiation100, heat_radiation80, heat_radiation60, heat_radiation40, " & _
                                      "oil_cooler100, oil_cooler80, oil_cooler60, oil_cooler40, link, mainheat100_u, mainheat80_u, mainheat60_u, mainheat40_u, " & _
                                      "lt_heat100_u, lt_heat80_u, lt_heat60_u, lt_heat40_u, oil_cooler100_u, oil_cooler80_u, oil_cooler60_u, oil_cooler40_u"

    Public Const ALL_MTU As String = "id, model, rpm, fuel, burn_type, nox, comp_ratio, " & _
                                  "Voltage, JW_in, JW_out, IC_in, IC_out, ElePow100, ElePow75, ElePow50, EngPow100, EngPow75, EngPow50, mainheat100, mainheat75,	mainheat50, " & _
                                  "LT_heat100, LT_heat75, LT_heat50, ExFlow100, ExFlow75, ExFlow50, ExTemp100, ExTemp75, ExTemp50, fuelcon100, fuelcon75, fuelcon50, link, heat_radiation100 as vent_heat, " & _
                                  "mainheat100_u, mainheat75_u,	mainheat50_u, LT_heat100_u, LT_heat75_u, LT_heat50_u, fuelcon100_u,	fuelcon75_u, fuelcon50_u"

    Public Const ALL_MAN As String = STANDARD_STATS & ", engpow100, engpow75, engpow50, elepow100, elepow75, elepow50, mainheat100, mainheat75, mainheat50, " & _
                                  "exflow100, exflow75, exflow50, extemp100, extemp75, extemp50, lt_heat100, lt_heat75, lt_heat50, " & _
                                  "fuelcon100, fuelcon75, fuelcon50, heat_radiation100, heat_radiation75, heat_radiation50, link, mainheat100_u, mainheat75_u, mainheat50_u"
#End Region

#Region "GENSET STRINGS"
    Public Sub FillGensetDGVCols(dgv As DataGridView)
        dgv.ColumnCount = 24
        dgv.Columns(0).Name = "Product ID" : dgv.Columns(1).Name = "MFR" : dgv.Columns(2).Name = "Model" : dgv.Columns(3).Name = "RPM"
        dgv.Columns(4).Name = "FuelType" : dgv.Columns(5).Name = "Engine_KW" : dgv.Columns(6).Name = "LT_Heat" : dgv.Columns(7).Name = "FuelCon"
        dgv.Columns(8).Name = "FuelCon_bHP" : dgv.Columns(9).Name = "Steam_Recov" : dgv.Columns(10).Name = "JW_to_PRM" : dgv.Columns(11).Name = "ExHeatRecov_PRM"
        dgv.Columns(12).Name = "OilCool_PRM" : dgv.Columns(13).Name = "Total_PRM" : dgv.Columns(14).Name = "IC_to_SEC" : dgv.Columns(15).Name = "Ele_Eff"
        dgv.Columns(16).Name = "Therm_Eff" : dgv.Columns(17).Name = "Total_Eff" : dgv.Columns(18).Name = "PW_Flow" : dgv.Columns(19).Name = "PW_In"
        dgv.Columns(20).Name = "PW_Out" : dgv.Columns(21).Name = "SW_Flow" : dgv.Columns(22).Name = "SW_In" : dgv.Columns(23).Name = "SW_Out"
    End Sub

    ' WHAT I ALREADY HAVE:
    ' id, mfr, model, rpm, fuel, burn_type, nox, elepow100
    Public Const GUASCOR_GENSET As String = "elepow80, elepow60, elepow40, jw_out, jw_flow, ic_in, ic_flow, exflow100, exflow80, exflow60, exflow40, extemp100, extemp80, extemp60, extemp40, " & _
                                        "mainheat100_u, mainheat80_u, mainheat60_u, mainheat40_u, fuelcon100_u, fuelcon80_u, fuelcon60_u, fuelcon40_u, oil_cooler100_u, oil_cooler80_u, oil_cooler60_u, oil_cooler40_u"

    Public Const MTU_GENSET As String = "elepow75, elepow50, engpow100, engpow75, engpow50, jw_in, jw_out, ic_in, ic_out, exflow100, exflow75, exflow50, " & _
                                     "extemp100, extemp75, extemp50, fuelcon100_u, fuelcon75_u, fuelcon50_u, mainheat100_u, mainheat75_u, mainheat50_u, lt_heat100_u, lt_heat75_u, lt_heat50_u"

    Public Const PAIRING_QUERY As String = "SELECT TOP(1) g.id FROM generators AS g, engines AS e WHERE e.id=@id AND e.rpm = g.rpm AND g.elepow > e.elepow100"
#End Region

#Region "MATH CONSTANTS"
    Public Const ELE_EFF_PERCENT As Double = 0.95
    Public Const ELE_CONVERSION As Double = 0.7457
#End Region
End Module