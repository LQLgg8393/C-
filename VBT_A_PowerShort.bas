Attribute VB_Name = "VBT_A_PowerShort"
Option Explicit
Private Const m_CAPTUREWAIT_dbl As Double = 0.01    ' new spec settling time in sec,

Public Function A_PowerShort_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_AllPowerPins_PL As PinList, _
       in_MeterRanges_str As String, _
       in_AveragePointCnt_lng As Long) As Long

' IMPORTANT NOTICE !!!
' this function assumes all the power pins are either UVS256 or HexVS pins, if you have DCVI power pins, please update this function!

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__Basic", "A_PowerShort_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim Site As Variant
    Dim i As Long
    Dim in_MeasurePinsCount_lng As Long
    Dim i_PinList_str() As String
    Dim in_AllPowerPins_str() As String

    Dim t_IDDQHexVS_PLD As New PinListData
    Dim t_IDDQVHDVS_PLD As New PinListData
    Dim t_IDDQUVS64_PLD As New PinListData
    Dim t_IDDQUVS256Ufp_PLD As New PinListData

    TheHdw.Digital.Pins("VDD_ODIO_BIAS").InitState = chInitLo
    
    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl  ' maybe this should always be tlUnpowered ?

    '   II. Setup P-Set
    ' 1. setup the lookup table
    Call TheDPS.GeneratePinTypeTable(in_AllPowerPins_PL)
    ' 2. setup meter mode to FVMI
    Call TheDPS.SetupDCVSFVMIPSets(in_AllPowerPins_PL, in_MeterRanges_str, i_PinList_str, "PowerShort")
    ' 3. set the I meter range before measurement
    Call TheDPS.ApplyDCVSFVMIPSets(in_AllPowerPins_PL, "PowerShort", m_CAPTUREWAIT_dbl)



    '   III. strobe but do not read back
    ' sadly HexVS does not have Hardware averaging
    TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Strobe in_AveragePointCnt_lng, 200000
    ' UVS256 has hardware averaging through filter (refer to IGXL Help)
    TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Filter.Bypass = False
    TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Filter.Value = 200000# / in_AveragePointCnt_lng
    TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Strobe 1, 200000#
    '
    'TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Strobe in_AveragePointCnt_lng, 200000  ' will confirm
    '
    TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Filter.Bypass = False
    TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Filter.Value = 200000# / in_AveragePointCnt_lng
    TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Strobe 1, 200000#
    
    TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Filter.Bypass = False
    TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Filter.Value = 200000# / in_AveragePointCnt_lng
    TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Strobe 1, 200000#
    '
    TheHdw.Wait in_AveragePointCnt_lng / 200000#


    '   IV. readback DCVS I measurement results
    Call TheExec.DataManager.DecomposePinList(in_AllPowerPins_PL, in_AllPowerPins_str, in_MeasurePinsCount_lng)
    '    On Error GoTo retestPS 'Retest per site if DCVSPin Data not found (error DCVS: 0042)
    '    TheExec.Error.Behavior("DCVS:0042") = tlErrorContinueOnError
    '
    'normaltest:
    On Error Resume Next            'Retest per site if DCVSPin Data not found (error DCVS: 0042)
    TheExec.Error.Behavior("DCVS:0042") = tlErrorContinueOnError

    t_IDDQVHDVS_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Read(tlNoStrobe)
    t_IDDQHexVS_PLD = TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Read(tlNoStrobe, in_AveragePointCnt_lng)
    t_IDDQUVS64_PLD = TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Read(tlStrobe) ', in_AveragePointCnt_lng)
    t_IDDQUVS256Ufp_PLD = TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Read(tlStrobe)
    i = 0
    While t_IDDQVHDVS_PLD.Pins.count + t_IDDQHexVS_PLD.Pins.count + t_IDDQUVS64_PLD.Pins.count + t_IDDQUVS256Ufp_PLD.Pins.count < in_MeasurePinsCount_lng
        DoEvents
        i = i + 1
        If i > 100 Then
            On Error GoTo errHandler
        End If
        For Each Site In TheExec.Sites
            t_IDDQVHDVS_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Read(tlNoStrobe)
            t_IDDQHexVS_PLD = TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Read(tlNoStrobe, in_AveragePointCnt_lng)
            t_IDDQUVS64_PLD = TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Read(tlStrobe) ', in_AveragePointCnt_lng)
            t_IDDQUVS256Ufp_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Read(tlStrobe) ', in_AveragePointCnt_lng)
        Next Site
    Wend

    On Error GoTo errHandler


    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019



    For i = 0 To UBound(i_PinList_str)
        If InStr(1, TheDPS.HexVSPins, i_PinList_str(i), vbTextCompare) Then
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQHexVS_PLD.Pins(i_PinList_str(i)), _
                    forceResults:=tlForceFlow, _
                    TName:=i_PinList_str(i)
        ElseIf InStr(1, TheDPS.UVS256Pins, i_PinList_str(i), vbTextCompare) Then
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQVHDVS_PLD.Pins(i_PinList_str(i)), _
                    forceResults:=tlForceFlow, _
                    TName:=i_PinList_str(i)
        ElseIf InStr(1, TheDPS.UVS64Pins, i_PinList_str(i), vbTextCompare) Then
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQUVS64_PLD.Pins(i_PinList_str(i)), _
                    forceResults:=tlForceFlow, _
                    TName:=i_PinList_str(i)
                    
        ElseIf InStr(1, TheDPS.UVS256UfpPins, i_PinList_str(i), vbTextCompare) Then
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQUVS256Ufp_PLD.Pins(i_PinList_str(i)), _
                    forceResults:=tlForceFlow, _
                    TName:=i_PinList_str(i)
        End If
    Next

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    
    TheHdw.Digital.Pins("VDD_ODIO_BIAS").InitState = chInitoff

    '   X.

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function




