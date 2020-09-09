Attribute VB_Name = "VBT_A_IDDQ"
Option Explicit

Private Const m_CAPTUREWAIT_dbl As Double = 0.0005    ' new spec settling time in sec,

'=================For DVS Module=============================================================
Public g_IDDQResult_PLD_BefDVS As New PinListData
Public g_IDDQResult_PLD_AftDVS As New PinListData

'Public g_Aft_DVS_Flag As Boolean

Public Function A_IDDQ_wflag_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_MeasurePatName_PT As Pattern, _
       in_StrobesPerPoint_lng As Long, _
       in_MeasurePointCnt_lng As Long, _
       in_MeasurePins_PL As PinList, _
       in_MeterRanges_str As String, _
       in_AftDVSFlag__bool As Boolean, _
       in_EfuseUpdate_bool As Boolean) As Long
       
    

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__IDDQ", "A_IDDQ_Delta_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long
    Dim j As Long
    Dim Site As Variant
    Dim in_MeasurePinsCount_lng As Long

    Dim i_PinList_str() As String
    Dim in_MeasurePins_str() As String

    Dim i_IDDQHexVS_PLD As New PinListData
    Dim i_IDDQVHDVS_PLD As New PinListData
    Dim i_IDDQUVS64_PLD As New PinListData
    Dim i_IDDQUVS256Ufp_PLD As New PinListData

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_IDDQResultsMin_PLD As New PinListData
    Dim t_IDDQResultsMax_PLD As New PinListData
    Dim t_IDDQResultsMean_PLD As New PinListData
    Dim t_IDDQResults_PLD As New PinListData

    Dim i_TempStr_str As String
    Dim CurInstanceName As String
    CurInstanceName = TheExec.DataManager.InstanceName

    ' For DVS  IDDQ bypass
    If Not (CurInstanceName = "DC_IDDQ_DVS1" Or DVS_EFUSE_ITEM.dvs_allsite_execute = False Or in_AftDVSFlag__bool = False) Then
        TheExec.Datalog.WriteComment "Skip This IDDQ Test: " & CurInstanceName & "!!!"
        Exit Function
    End If

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl


    '   II. Setup P-Set
    'Call TheDPS.GeneratePinTypeTable(in_MeasurePins_PL)

    Call TheDPS.SetupDCVSFVMIPSets(in_MeasurePins_PL, in_MeterRanges_str, i_PinList_str, "User")
    ' UVS256 has hardware averaging through filter
    If TheDPS.UVS256Pins <> "" Then
        TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Filter.Bypass = False
        TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Filter.Value = 200000# / in_StrobesPerPoint_lng
    End If
    
    If TheDPS.UVS64Pins <> "" Then
        TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Filter.Bypass = False
        TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Filter.Value = 200000# / in_StrobesPerPoint_lng
    End If
    
    If TheDPS.UVS256UfpPins <> "" Then
        TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Filter.Bypass = False
        TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Filter.Value = 200000# / in_StrobesPerPoint_lng
    End If

    '   III. Looping the test
    TheHdw.Patterns(in_MeasurePatName_PT).Start
    For i = 1 To in_MeasurePointCnt_lng
        ' 1. stop at measurement point
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

        ' 2. set the I meter range before measurement
        Call TheDPS.ApplyDCVSFVMIPSets(in_MeasurePins_PL, "User", m_CAPTUREWAIT_dbl)

        ' 3. strobe but do not read back
        ' sadly HexVS does not have Hardware averaging
        If TheDPS.HexVSPins <> "" Then _
           TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Strobe in_StrobesPerPoint_lng
        ' UVS256 has hardware averaging through filter
        If TheDPS.UVS256Pins <> "" Then _
           TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Strobe 1, 200000# / in_StrobesPerPoint_lng
           
        If TheDPS.UVS64Pins <> "" Then _
           TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Strobe 1, 200000# / in_StrobesPerPoint_lng
           
        If TheDPS.UVS256UfpPins <> "" Then _
           TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Strobe 1, 200000# / in_StrobesPerPoint_lng
           
        TheHdw.Wait in_StrobesPerPoint_lng / 200000#

        ' 4. restore I meter range after measurement
        Call TheDPS.ApplyDCVSFVMIPSets(in_MeasurePins_PL, "Default", 0.0001)

        ' 5. continue to the next point
        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    ' restore
    TheHdw.Digital.Patgen.HaltWait

    '   IV. Retrieve Test result
    ' retrieve pattern burst result
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    ' readback DCVS I measurement results

    'Retrive Measure pins count
    Call TheExec.DataManager.DecomposePinList(in_MeasurePins_PL, in_MeasurePins_str, in_MeasurePinsCount_lng)

    On Error Resume Next            'Retest per site if DCVSPin Data not found (error DCVS: 0042)
    TheExec.Error.Behavior("DCVS:0042") = tlErrorContinueOnError
    
    If TheDPS.UVS256Pins <> "" Then i_IDDQVHDVS_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Read(tlNoStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
    If TheDPS.UVS256UfpPins <> "" Then i_IDDQVHDVS_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Read(tlNoStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
    If TheDPS.UVS64Pins <> "" Then i_IDDQUVS64_PLD = TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Read(tlNoStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
    If TheDPS.HexVSPins <> "" Then i_IDDQHexVS_PLD = TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Read(tlNoStrobe, in_MeasurePointCnt_lng * in_StrobesPerPoint_lng, , tlDCVSMeterReadingFormatArray)
    i = 0
    While i_IDDQVHDVS_PLD.Pins.count + i_IDDQHexVS_PLD.Pins.count + i_IDDQUVS64_PLD.Pins.count + i_IDDQUVS256Ufp_PLD.Pins.count < in_MeasurePinsCount_lng
        DoEvents
        i = i + 1
        If i > 100 Then
            On Error GoTo errHandler
        End If
        For Each Site In TheExec.Sites
            i_IDDQVHDVS_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256Pins).Meter.Read(tlStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
            i_IDDQUVS64_PLD = TheHdw.DCVS.Pins(TheDPS.UVS64Pins).Meter.Read(tlStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
            i_IDDQUVS256Ufp_PLD = TheHdw.DCVS.Pins(TheDPS.UVS256UfpPins).Meter.Read(tlStrobe, in_MeasurePointCnt_lng, , tlDCVSMeterReadingFormatArray)
            i_IDDQHexVS_PLD = TheHdw.DCVS.Pins(TheDPS.HexVSPins).Meter.Read(tlStrobe, in_MeasurePointCnt_lng * in_StrobesPerPoint_lng, , tlDCVSMeterReadingFormatArray)
        Next Site
    Wend

    On Error GoTo errHandler


    If TheExec.TesterMode = testModeOffline Then
    
        Dim tmp_dspwave As New DSPWave
        Dim tmp_arr() As Double
        ReDim tmp_arr(in_MeasurePointCnt_lng * in_StrobesPerPoint_lng - 1)
        tmp_dspwave.CreateRandom 99, 999, in_MeasurePointCnt_lng * in_StrobesPerPoint_lng, DspDouble
        tmp_arr() = tmp_dspwave.Data
        
        For i = 0 To i_IDDQHexVS_PLD.Pins.count - 1
            i_IDDQHexVS_PLD.Pins(i) = tmp_arr()
        Next i
    End If

    Call CalculateIDDQ(in_MeasurePointCnt_lng, in_StrobesPerPoint_lng, i_IDDQVHDVS_PLD, i_IDDQHexVS_PLD, i_IDDQUVS64_PLD, i_IDDQUVS256Ufp_PLD, t_IDDQResultsMin_PLD, t_IDDQResultsMax_PLD, t_IDDQResultsMean_PLD, t_IDDQResults_PLD)

    '   V. Datalog

    'TheExec.Flow.Limits.key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit _
            ResultVal:=t_FuncResult_SBOOL, _
            TName:="IDDQ_FUNCTION_RESULT", _
            forceResults:=tlForceFlow

    For i = 0 To t_IDDQResultsMean_PLD.Pins.count - 1
        If in_MeasurePointCnt_lng = 1 Then
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQResultsMean_PLD.Pins(i_PinList_str(i)), _
                    PinName:=i_PinList_str(i), _
                    TName:=i_PinList_str(i) + "_MEAN", _
                    forceResults:=tlForceFlow
        Else
''            For Each Site In TheExec.Sites
''                For j = 0 To in_MeasurePointCnt_lng - 1
''                    TheExec.Flow.TestLimit _
''                        ResultVal:=t_IDDQResults_PLD.Pins(i_PinList_str(i))(Site).Element(j), _
''                        TName:=i_PinList_str(i) + "@index" + CStr(j), _
''                        forceResults:=tlForceFlow
''                Next j
''            Next Site
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQResultsMean_PLD.Pins(i_PinList_str(i)), _
                    PinName:=i_PinList_str(i), _
                    TName:=i_PinList_str(i) + "_MEAN", _
                    forceResults:=tlForceFlow
            ' TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQResultsMin_PLD.Pins(i_PinList_str(i)), _
                    PinName:=i_PinList_str(i), _
                    TName:=i_PinList_str(i) + "_MIN", _
                    forceResults:=tlForceFlow
            ' TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            TheExec.Flow.TestLimit _
                    ResultVal:=t_IDDQResultsMax_PLD.Pins(i_PinList_str(i)), _
                    PinName:=i_PinList_str(i), _
                    TName:=i_PinList_str(i) + "_MAX", _
                    forceResults:=tlForceFlow
        End If
    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    
    
    ' ================================================================================
    '                Update Efuse Table and Trim Data Init
    ' ================================================================================
    If in_EfuseUpdate_bool = True Then
        For Each Site In TheExec.Sites
            If TheExec.TesterMode = testModeOnline Then
                Call hiefuse.MULT_EFUSE_ITEM_UPDATE("IDDQ_CPUB", 0, Site, CDbl(t_IDDQResultsMean_PLD.Pins("VDD08_CPU_BM")(Site)) * 1000, 0, 0)
                Call hiefuse.MULT_EFUSE_ITEM_UPDATE("IDDQ_CPUL", 0, Site, CDbl(t_IDDQResultsMean_PLD.Pins("VDD08_CPU_L")(Site)) * 1000, 0, 0)
                Call hiefuse.MULT_EFUSE_ITEM_UPDATE("IDDQ_GPU", 0, Site, CDbl(t_IDDQResultsMean_PLD.Pins("VDD07_GPU")(Site)) * 1000, 0, 0)
            ElseIf TheExec.TesterMode = testModeOffline Then
            End If
            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("IDDQ_CPUB")).gDataValue(Site) = CDbl(t_IDDQResultsMean_PLD.Pins("VDD08_CPU_BM")(Site)) * 1000
            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("IDDQ_CPUL")).gDataValue(Site) = CDbl(t_IDDQResultsMean_PLD.Pins("VDD08_CPU_L")(Site)) * 1000
            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("IDDQ_GPU")).gDataValue(Site) = CDbl(t_IDDQResultsMean_PLD.Pins("VDD07_GPU")(Site)) * 1000
        
        Next Site
    End If
   
    '   X.
    ' ================================================================================
    '                        Delta IDDQ
    ' ================================================================================

    'Only IDDQ_DVS can update DVS global variant, Jeremy
    If CurInstanceName = "DC_IDDQ_DVS1" Or CurInstanceName = "DC_IDDQ_DVS2" Then
        If in_AftDVSFlag__bool = False Then

            g_IDDQResult_PLD_BefDVS = t_IDDQResultsMean_PLD       ' Restore IDDQ result before DVS

            i_TempStr_str = i_PinList_str(0)
            For i = 1 To in_MeasurePinsCount_lng - 1
                i_TempStr_str = i_TempStr_str + "," + i_PinList_str(i)
            Next i
            'g_PinsBefDVS_CSL.value = i_TempStr_str

        Else

            g_IDDQResult_PLD_AftDVS = t_IDDQResultsMean_PLD      ' Restore IDDQ result after DVS

            i_TempStr_str = i_PinList_str(0)
            For i = 1 To in_MeasurePinsCount_lng - 1
                i_TempStr_str = i_TempStr_str + "," + i_PinList_str(i)
            Next i
            'g_PinsAftDVS_CSL.value = i_TempStr_str

        End If
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function CalculateIDDQ( _
        ByVal in_MeasurePointCnt_lng As Long, _
        ByVal in_StrobesPerPoint_lng As Long, _
        ByRef io_IDDQVHDVS_PLD As PinListData, _
        ByRef io_IDDQHexVS_PLD As PinListData, _
        ByRef io_IDDQUVS64_PLD As PinListData, _
        ByRef io_IDDQUVS256Ufp_PLD As PinListData, _
        ByRef io_IDDQResultsMin_PLD As PinListData, _
        ByRef io_IDDQResultsMax_PLD As PinListData, _
        ByRef io_IDDQResultsMean_PLD As PinListData, _
        ByRef io_IDDQResults_PLD As PinListData _
      ) As Long
' only call this function when reading from array!!!

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT_D01_IDDQ", "CalculateIDDQ")

    Dim Site As Variant
    Dim i As Long
    Dim i_PinData_PD As Variant
    Dim i_tmp_DSP As New DSPWave
    Dim i_tmp_HexVS_DSP As New DSPWave
    Dim i_tmp_UVS64_DSP As New DSPWave
    Dim i_tmp_UVS256Ufp_DSP As New DSPWave
    Dim i_array_dbl() As Double
    Dim i_max_dbl As Double
    Dim i_min_dbl As Double
    Dim i_mean_dbl As Double


    Set io_IDDQResultsMin_PLD = New PinListData
    Set io_IDDQResultsMax_PLD = New PinListData
    Set io_IDDQResultsMean_PLD = New PinListData
    Set io_IDDQResults_PLD = New PinListData

    For Each i_PinData_PD In io_IDDQVHDVS_PLD.Pins
        io_IDDQResultsMin_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMax_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMean_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResults_PLD.AddPin (i_PinData_PD.Name)
        For Each Site In TheExec.Sites
            i_tmp_DSP.Data = i_PinData_PD.Value
'        Next Site
'        'io_IDDQVHDVS_PLD.Pins(i_PinData_PD) = i_tmp_DSP
'        For Each Site In TheExec.Sites
            io_IDDQResultsMin_PLD.Pins(i_PinData_PD) = i_tmp_DSP.CalcMinimumValue
            io_IDDQResultsMax_PLD.Pins(i_PinData_PD) = i_tmp_DSP.CalcMaximumValue
            io_IDDQResultsMean_PLD.Pins(i_PinData_PD) = i_tmp_DSP.CalcMean
            io_IDDQResults_PLD.Pins(i_PinData_PD) = i_tmp_DSP
        Next Site
    Next
    
    i_tmp_HexVS_DSP.CreateConstant 999#, in_MeasurePointCnt_lng, DspDouble
    
    For Each i_PinData_PD In io_IDDQHexVS_PLD.Pins
        io_IDDQResultsMin_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMax_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMean_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResults_PLD.AddPin (i_PinData_PD.Name)
        For Each Site In TheExec.Sites
            i_tmp_DSP.Data = i_PinData_PD.Value

            For i = 0 To in_MeasurePointCnt_lng - 1
                i_tmp_HexVS_DSP.Element(i) = i_tmp_DSP.Select(i * in_StrobesPerPoint_lng, 1, in_StrobesPerPoint_lng).CalcMean
            Next i
            io_IDDQResultsMin_PLD.Pins(i_PinData_PD) = i_tmp_HexVS_DSP.CalcMinimumValue
            io_IDDQResultsMax_PLD.Pins(i_PinData_PD) = i_tmp_HexVS_DSP.CalcMaximumValue
            io_IDDQResultsMean_PLD.Pins(i_PinData_PD) = i_tmp_HexVS_DSP.CalcMean
            io_IDDQResults_PLD.Pins(i_PinData_PD) = i_tmp_HexVS_DSP
        Next Site
    Next
    
    i_tmp_UVS64_DSP.CreateConstant 999#, in_MeasurePointCnt_lng, DspDouble
    
    For Each i_PinData_PD In io_IDDQUVS64_PLD.Pins
        io_IDDQResultsMin_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMax_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMean_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResults_PLD.AddPin (i_PinData_PD.Name)
        For Each Site In TheExec.Sites
            i_tmp_UVS64_DSP.Data = i_PinData_PD.Value
'        Next Site
'        'io_IDDQVHDVS_PLD.Pins(i_PinData_PD) = i_tmp_DSP
'        For Each Site In TheExec.Sites
            io_IDDQResultsMin_PLD.Pins(i_PinData_PD) = i_tmp_UVS64_DSP.CalcMinimumValue
            io_IDDQResultsMax_PLD.Pins(i_PinData_PD) = i_tmp_UVS64_DSP.CalcMaximumValue
            io_IDDQResultsMean_PLD.Pins(i_PinData_PD) = i_tmp_UVS64_DSP.CalcMean
            io_IDDQResults_PLD.Pins(i_PinData_PD) = i_tmp_UVS64_DSP
        Next Site
    Next
    
    
    i_tmp_UVS256Ufp_DSP.CreateConstant 999#, in_MeasurePointCnt_lng, DspDouble
    
    For Each i_PinData_PD In io_IDDQUVS256Ufp_PLD.Pins
        io_IDDQResultsMin_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMax_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResultsMean_PLD.AddPin (i_PinData_PD.Name)
        io_IDDQResults_PLD.AddPin (i_PinData_PD.Name)
        For Each Site In TheExec.Sites
            i_tmp_UVS256Ufp_DSP.Data = i_PinData_PD.Value
'        Next Site
'        'io_IDDQVHDVS_PLD.Pins(i_PinData_PD) = i_tmp_DSP
'        For Each Site In TheExec.Sites
            io_IDDQResultsMin_PLD.Pins(i_PinData_PD) = i_tmp_UVS256Ufp_DSP.CalcMinimumValue
            io_IDDQResultsMax_PLD.Pins(i_PinData_PD) = i_tmp_UVS256Ufp_DSP.CalcMaximumValue
            io_IDDQResultsMean_PLD.Pins(i_PinData_PD) = i_tmp_UVS256Ufp_DSP.CalcMean
            io_IDDQResults_PLD.Pins(i_PinData_PD) = i_tmp_UVS256Ufp_DSP
        Next Site
    Next
    

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function A_IDDQ_Delta_vbt( _
       in_MeasPins_PL As PinList) As Long
'
    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__IDDQ", "A_IDDQ_Delta_vbt", TheExec.DataManager.InstanceName)
    If TheExec.Sites.Active.count = 0 Then Exit Function

    ' ================================================================================
    '                        Declare variables
    ' ================================================================================
    Dim Site

    Dim i_IDDQ_Delta_PLD As New PinListData
    Dim i_MeasPins_str() As String
    Dim i_PinNum_lng As Long

    Dim i_PinIndex_lng As Variant

    ' ================================================================================
    '                        Initialize Settings
    ' ================================================================================
    'Nop by Jeremy to suit for 93K flow
    'Call IDDQ_Delta_Check(in_MeasPins_PL)

    Call TheExec.DataManager.DecomposePinList(in_MeasPins_PL, i_MeasPins_str(), i_PinNum_lng)

    Set i_IDDQ_Delta_PLD = Nothing

    For i_PinIndex_lng = 0 To i_PinNum_lng - 1
        i_IDDQ_Delta_PLD.AddPin (i_MeasPins_str(i_PinIndex_lng))
        i_IDDQ_Delta_PLD.Pins(i_MeasPins_str(i_PinIndex_lng)).Value = 0.1 * 10 ^ -9
    Next i_PinIndex_lng

    ' ================================================================================
    '                        Calculate and datalog
    ' ================================================================================
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName
    For Each Site In TheExec.Sites
        If DVS_EFUSE_ITEM.DVS_FLAG_RD(Site) = 0 And DVS_EFUSE_ITEM.DVS_RESULT_RD(Site) = 0 Then
            For i_PinIndex_lng = 0 To i_PinNum_lng - 1

                TheExec.Flow.TestLimitIndex = i_PinIndex_lng + 1
                i_IDDQ_Delta_PLD.Pins(i_MeasPins_str(i_PinIndex_lng)) = g_IDDQResult_PLD_AftDVS.Pins(i_MeasPins_str(i_PinIndex_lng)).Divide(g_IDDQResult_PLD_BefDVS.Pins(i_MeasPins_str(i_PinIndex_lng)))
                TheExec.Flow.TestLimit ResultVal:=i_IDDQ_Delta_PLD.Pins(i_MeasPins_str(i_PinIndex_lng)), TName:=i_MeasPins_str(i_PinIndex_lng), forceResults:=tlForceFlow

            Next i_PinIndex_lng
        End If

        If DVS_EFUSE_ITEM.DVS_FLAG_RD(Site) = 1 And DVS_EFUSE_ITEM.DVS_RESULT_RD(Site) = 0 Then
            'TheExec.Flow.TestLimitIndex = 23    '0
            TheExec.Flow.TestLimit 0, forceResults:=tlForceFlow, TName:="PREVIOUS_DELTAIDDQ_FAIL"
        End If
        If DVS_EFUSE_ITEM.DVS_FLAG_RD(Site) = 1 And DVS_EFUSE_ITEM.DVS_RESULT_RD(Site) = 1 Then
            'TheExec.Flow.TestLimitIndex = 24    '0
            TheExec.Flow.TestLimit 1, forceResults:=tlForceFlow, TName:="PREVIOUS_DELTAIDDQ_PASS"
        End If
    Next Site
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function







