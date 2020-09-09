Attribute VB_Name = "VBT_A_HPM"
Option Explicit

Global RegNameArr() As String
Global ValidNameArr() As String

Global ValidIndex_dsp As New DSPWave
Global RegIndex_dsp As New DSPWave

Private Const m_SETTLEWAIT_dbl As Double = 0.01   ' new spec settling time in sec,


Public Function HPM_ReadCode_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_CaptureSampleSize_lng As Long, _
       in_RegisterNames_str As String, _
       in_ValidRegisterNames_str As String, _
       in_PowerPins_PL As PinList, _
       in_TestConditionList_str As String, _
        in_TrimUpdate_bool As Boolean) As Long
' EDITFORMAT1 1,RelayMode,tlRelayMode,,UnPowered will reset the DUT,in_RelayMode_tl|2,ConfigPatName_PT,Pattern,,,in_ConfigPatName_PT|3,CapturePatName_PT,Pattern,,,in_CapturePatName_PT|4,CaptureSampleSize_lng,Long,,,in_CaptureSampleSize_lng|5,RegisterNames_str,String,,comma separated list,in_RegisterNames_str|6,ValidRegisterNames_str,String,,comma separated list,in_ValidRegisterNames_str|7,PowerPins_PL,PinList,,,in_PowerPins_PL|8,TestConditionList_str,String,,comma separated list,in_TestConditionList_str

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__HPM", "HPM_ReadCode_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long
    Dim j As Long
    Dim Site As Variant

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_SignalName_str As String = "HPM_ReadCode"

    Dim i_RunPrecondition_bool As Boolean
    Dim i_CaptureValid_bool As Boolean

    Const i_RegisterBitWidth_lng As Long = 10
    Const i_ValidRegisterBitWidth_lng As Long = 1

    Dim i_TestCount_lng As Long
    Dim i_TestCondition_str() As String

    Dim i_BaseVoltage_dbl As Double
    Dim i_TestVoltage_dbl() As Double

    Dim i_RegName_str() As String
    Dim i_ValidRegName_str() As String

    Dim i_RegData_PLD As New PinListData
    Dim i_ValidRegData_PLD As New PinListData
    Dim t_RegData_DSP As New DSPWave
    Dim t_ValidRegData_DSP As New DSPWave


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    in_RegisterNames_str = Trim(in_RegisterNames_str)
    i_RegName_str = Split(in_RegisterNames_str, ",")
    in_ValidRegisterNames_str = Trim(in_ValidRegisterNames_str)
    If in_ValidRegisterNames_str = "" Then
        i_CaptureValid_bool = False
    Else
        i_CaptureValid_bool = True
        i_ValidRegName_str = Split(in_ValidRegisterNames_str, ",")
    End If

    i_TestCondition_str = Split(in_TestConditionList_str, ",")
    i_TestCount_lng = UBound(i_TestCondition_str)
    ReDim i_TestVoltage_dbl(i_TestCount_lng)

    i_BaseVoltage_dbl = TheHdw.DCVS.Pins(in_PowerPins_PL).voltage.Main
    ' if an error reported here, it is possible that 'in_PowerPins_PL' _
      contains more than one Power Pin. Correct this typo in the test instance sheet, _
      if you continue from here, zero voltage will be applied for all steps.
    For i = 0 To i_TestCount_lng
        If (i_TestCondition_str(i) = "HV") Then
            i_TestVoltage_dbl(i) = 0.95
        ElseIf (i_TestCondition_str(i) = "MV") Then
            i_TestVoltage_dbl(i) = 0.8
        ElseIf (i_TestCondition_str(i) = "LV") Then
            i_TestVoltage_dbl(i) = 0.65
        Else
            TheExec.AddOutput "Cannot Determin Test Voltage, check Test Instance!", vbRed, True
            i_TestVoltage_dbl(i) = i_BaseVoltage_dbl
            Stop    ' Please correct this typo in the test instance sheet, _
                    if you continue from here, original voltage will be applied for this step
        End If
    Next i


    '   II. Setup DSSC
    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng, _
         i_CapCodes_PLD)


    '   III. Looping the test
    For i = 0 To i_TestCount_lng

        TheHdw.DCVS.Pins(in_PowerPins_PL).voltage.Main = i_TestVoltage_dbl(i)
        TheHdw.Wait m_SETTLEWAIT_dbl

        If i_RunPrecondition_bool Then
            TheHdw.Patterns(in_ConfigPatName_PT).Start
            TheHdw.Digital.Patgen.HaltWait
        End If
    
        If TheExec.TesterMode = testModeOnline Then
            TheHdw.Patterns(in_CapturePatName_PT).Start
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next i
    ' restore power voltage
    TheHdw.DCVS.Pins(in_PowerPins_PL).voltage.Main = i_BaseVoltage_dbl

    '   IV. Backgroud Calculation
    rundsp.E_01_HPM_Cal i_CapCodes_PLD, i_TestCount_lng, _
                        i_RegisterBitWidth_lng, i_RegData_PLD, i_CaptureValid_bool, _
                        i_ValidRegisterBitWidth_lng, i_ValidRegData_PLD

    t_RegData_DSP = i_RegData_PLD.Pins(i_CapturePins_PL)
    t_ValidRegData_DSP = i_ValidRegData_PLD.Pins(i_CapturePins_PL)

    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    For i = 0 To i_TestCount_lng
        
        For j = 0 To in_CaptureSampleSize_lng - 1
            TheExec.Flow.TestLimit _
                    ResultVal:=t_RegData_DSP.Element(j + (in_CaptureSampleSize_lng * i)), _
                    TName:=i_RegName_str(j) + "_" + i_TestCondition_str(i), _
                    PinName:=i_CapturePins_PL, _
                    forceResults:=tlForceFlow
            ' ================================================================================
            '                 Trim Data Init
            ' ================================================================================
            If in_TrimUpdate_bool = True And TheExec.DataManager.InstanceName = "HPM_DDR_300H8LVT" And i_TestCondition_str(i) = "MV" Then
                If i_RegName_str(j) = "DMC_HIDDRPHY_INTF_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("DMC_HPM_DDR_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (in_CaptureSampleSize_lng * i)))
            End If
        Next j
        If i_CaptureValid_bool Then
            For j = 0 To in_CaptureSampleSize_lng - 1
                TheExec.Flow.TestLimit _
                        ResultVal:=t_ValidRegData_DSP.Element(j + (in_CaptureSampleSize_lng * i)), _
                        TName:=i_ValidRegName_str(j) + "_" + i_TestCondition_str(i), _
                        PinName:=i_CapturePins_PL, _
                        forceResults:=tlForceFlow
            Next j
        End If
    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    '   X.

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function HPM_ReadCodeSpec_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_CaptureSampleSize_lng As Long, _
       in_SpecName_str As String, _
       in_TestConditionList_str As String, _
       in_TrimUpdate_bool As Boolean) As Long
' EDITFORMAT1 1,RelayMode,tlRelayMode,,UnPowered will reset the DUT,in_RelayMode_tl|2,ConfigPatName_PT,Pattern,,,in_ConfigPatName_PT|3,CapturePatName_PT,Pattern,,,in_CapturePatName_PT|4,CaptureSampleSize_lng,Long,,,in_CaptureSampleSize_lng|5,RegisterNames_str,String,,comma separated list,in_RegisterNames_str|6,ValidRegisterNames_str,String,,comma separated list,in_ValidRegisterNames_str|7,SpecName_str,String,,,in_SpecName_str|8,TestConditionList_str,String,,comma separated list,in_TestConditionList_str

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__HPM", "HPM_ReadCodeSpec_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long
    Dim j As Long
    Dim Site As Variant

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_SignalName_str As String = "HPM_ReadCode"

    Dim i_RunPrecondition_bool As Boolean

    Const i_RegisterBitWidth_lng As Long = 10
    Const i_ValidRegisterBitWidth_lng As Long = 1

    Dim i_TestCount_lng As Long
    Dim i_TestCondition_str() As String

    Dim i_BaseSpec_SDBL As New SiteDouble
    Dim i_TestSpec_dbl() As Double


    Dim i_RegData_PLD As New PinListData
    Dim i_ValidRegData_PLD As New PinListData
    Dim t_RegData_DSP As New DSPWave
    Dim t_ValidRegData_DSP As New DSPWave


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    i_TestCondition_str = Split(in_TestConditionList_str, ",")
    i_TestCount_lng = UBound(i_TestCondition_str)
    ReDim i_TestSpec_dbl(i_TestCount_lng)

    For i = 0 To i_TestCount_lng
        i_TestSpec_dbl(i) = LookupSpecTable("HPM_" + i_TestCondition_str(i))
    Next i

    '   II. prepare the Test Conditions
    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng, _
         i_CapCodes_PLD)


    '   III. Looping the test
    For i = 0 To i_TestCount_lng
        'ApplyUniformSpecToHW
        TheExec.Overlays.ApplyUniformSpecToHW in_SpecName_str, i_TestSpec_dbl(i), True, True
        TheHdw.Wait m_SETTLEWAIT_dbl

        If i_RunPrecondition_bool Then
            TheHdw.Patterns(in_ConfigPatName_PT).Start
            TheHdw.Digital.Patgen.HaltWait
        End If

        If TheExec.TesterMode = testModeOnline Then
            TheHdw.Patterns(in_CapturePatName_PT).Start
            TheHdw.Digital.Patgen.HaltWait
        End If

    Next i
    ' restore power voltage
    TheExec.Overlays.ApplyUniformSpecToHW in_SpecName_str, 1#, True, True
    
    'init hpm array
    If RegIndex_dsp.SampleSize = 0 Or ValidIndex_dsp.SampleSize = 0 Then Call Init_HPM_Array
    
    '   IV. Backgroud Calculation
    rundsp.E_01_HPM_Cal_All i_CapCodes_PLD, i_TestCount_lng, RegIndex_dsp, ValidIndex_dsp, _
                             i_RegisterBitWidth_lng, i_RegData_PLD, _
                            i_ValidRegisterBitWidth_lng, i_ValidRegData_PLD

    t_RegData_DSP = i_RegData_PLD.Pins(i_CapturePins_PL)
    t_ValidRegData_DSP = i_ValidRegData_PLD.Pins(i_CapturePins_PL)
    
    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    For i = 0 To i_TestCount_lng
        For j = 0 To RegIndex_dsp.SampleSize - 1
        
            If TheExec.DataManager.InstanceName = "HPM_ULVT" And RegNameArr(j) = "ENYO_U2_HPM_PC_ORG" And i_TestCondition_str(i) = "MV" Then
                gHPM_ULVT__enyo_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
            End If
            If TheExec.DataManager.InstanceName = "HPM_LVT" And i_TestCondition_str(i) = "MV" Then
                If RegNameArr(j) = "ENYO_MIDCORE_U2_HPM_PC_ORG" Then gHPM_LVT__enyo_midcore_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                If RegNameArr(j) = "ANANKE_U2_HPM_PC_ORG" Then gHPM_LVT__ananke_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                If RegNameArr(j) = "FCM_U2_HPM_PC_ORG" Then gHPM_LVT__FCM_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                If RegNameArr(j) = "G3D_U2_HPM_PC_ORG" Then gHPM_LVT__G3D_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                If RegNameArr(j) = "NPU_WRAP_U2_HPM_PC_ORG" Then gHPM_LVT__NPU_WRAP_u2_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                If RegNameArr(j) = "PERI_U1_HPM_PC_ORG" Then gHPM_LVT__PERI_u1_hpm_pc_org_MV = t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i))
                ' ================================================================================
                '              Trim Data Init
                ' ================================================================================
                If in_TrimUpdate_bool = True Then
                    If RegNameArr(j) = "FCM_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("FCM_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "G3D_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("G3D_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "MODEM5G_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("MODEM5G_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "MODEM_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("MODEM_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "NPU_NPU_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("NPU_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "PERI_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("PERI_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "ANANKE_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("ANANKE_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "ENYO_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("ENYO_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "TRYM_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("TRYM_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "AO_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("AO_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                    If RegNameArr(j) = "BBP5G_U1_HPM_PC_ORG" Then Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("BBP5G_HPM_LVT")).gDataValue = CDbl(t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)))
                End If
            End If

            TheExec.Flow.TestLimit _
                    ResultVal:=t_RegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)), _
                    TName:=RegNameArr(j) + "_" + i_TestCondition_str(i), _
                    PinName:=i_CapturePins_PL, _
                    forceResults:=tlForceFlow
            TheExec.Flow.TestLimit _
                    ResultVal:=t_ValidRegData_DSP.Element(j + (RegIndex_dsp.SampleSize * i)), _
                    TName:=ValidNameArr(j) + "_" + i_TestCondition_str(i), _
                    PinName:=i_CapturePins_PL, _
                    forceResults:=tlForceFlow
        Next j
    Next i
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    '   X.

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Sub Init_HPM_Array()


ReDim RegNameArr(49)

ReDim ValidNameArr(49)

Dim RegIndexArr() As Long
ReDim RegIndexArr(49)

Dim ValidIndexArr() As Long
ReDim ValidIndexArr(49)
ValidNameArr(0) = "AO_U1_HPM_HPM_PC_VALID"
ValidNameArr(1) = "AO_U2_HPM_HPM_PC_VALID"
ValidNameArr(2) = "FCM_U1_HPM_HPM_PC_VALID"
ValidNameArr(3) = "FCM_U2_HPM_HPM_PC_VALID"
ValidNameArr(4) = "G3D_U1_HPM_HPM_PC_VALID"
ValidNameArr(5) = "G3D_U2_HPM_HPM_PC_VALID"
ValidNameArr(6) = "MODEM5G_U1_HPM_HPM_PC_VALID"
ValidNameArr(7) = "MODEM5G_U2_HPM_HPM_PC_VALID"
ValidNameArr(8) = "MODEM_U1_HPM_HPM_PC_VALID"
ValidNameArr(9) = "MODEM_U2_HPM_HPM_PC_VALID"
ValidNameArr(10) = "NPU_NPU_U1_HPM_HPM_PC_VALID"
ValidNameArr(11) = "NPU_WRAP_U1_HPM_HPM_PC_VALID"
ValidNameArr(12) = "NPU_NPU_U2_HPM_HPM_PC_VALID"
ValidNameArr(13) = "NPU_WRAP_U2_HPM_HPM_PC_VALID"
ValidNameArr(14) = "PERI_U1_HPM_HPM_PC_VALID"
ValidNameArr(15) = "PERI_U2_HPM_HPM_PC_VALID"
ValidNameArr(16) = "ANANKE_U1_HPM_HPM_PC_VALID"
ValidNameArr(17) = "ANANKE_U2_HPM_HPM_PC_VALID"
ValidNameArr(18) = "ANANKE_U1_HPM_1_HPM_PC_VALID"
ValidNameArr(19) = "ANANKE_U2_HPM_1_HPM_PC_VALID"
ValidNameArr(20) = "ANANKE_U1_HPM_2_HPM_PC_VALID"
ValidNameArr(21) = "ANANKE_U2_HPM_2_HPM_PC_VALID"
ValidNameArr(22) = "ANANKE_U1_HPM_3_HPM_PC_VALID"
ValidNameArr(23) = "ANANKE_U2_HPM_3_HPM_PC_VALID"
ValidNameArr(24) = "BBP5G_U1_HPM_PC_VALID"
ValidNameArr(25) = "BBP5G_U2_HPM_PC_VALID"
ValidNameArr(26) = "ENYO_U1_HPM_HPM_PC_VALID"
ValidNameArr(27) = "ENYO_U2_HPM_HPM_PC_VALID"
ValidNameArr(28) = "ENYO_MIDCORE_U1_HPM_HPM_PC_VALID"
ValidNameArr(29) = "ENYO_MIDCORE_U2_HPM_HPM_PC_VALID"
ValidNameArr(30) = "ENYO_MIDCORE_U1_HPM_1_HPM_PC_VALID"
ValidNameArr(31) = "ENYO_MIDCORE_U2_HPM_1_HPM_PC_VALID"
ValidNameArr(32) = "ENYO_MIDCORE_U1_HPM_2_HPM_PC_VALID"
ValidNameArr(33) = "ENYO_MIDCORE_U2_HPM_2_HPM_PC_VALID"
ValidNameArr(34) = "TRYM_U1_HPM_PC_VALID"
ValidNameArr(35) = "TRYM_U2_HPM_PC_VALID"
ValidNameArr(36) = "TRYM_U1_HPM_1_PC_VALID"
ValidNameArr(37) = "TRYM_U2_HPM_1_PC_VALID"
ValidNameArr(38) = "TRYM_U1_HPM_2_PC_VALID"
ValidNameArr(39) = "TRYM_U2_HPM_2_PC_VALID"
ValidNameArr(40) = "TRYM_U1_HPM_3_PC_VALID"
ValidNameArr(41) = "TRYM_U2_HPM_3_PC_VALID"
ValidNameArr(42) = "TRYM_U1_HPM_4_PC_VALID"
ValidNameArr(43) = "TRYM_U2_HPM_4_PC_VALID"
ValidNameArr(44) = "TRYM_U1_HPM_5_PC_VALID"
ValidNameArr(45) = "TRYM_U2_HPM_5_PC_VALID"
ValidNameArr(46) = "TRYM_U1_HPM_6_PC_VALID"
ValidNameArr(47) = "TRYM_U2_HPM_6_PC_VALID"
ValidNameArr(48) = "TRYM_U1_HPM_7_PC_VALID"
ValidNameArr(49) = "TRYM_U2_HPM_7_PC_VALID"


ValidIndexArr(0) = 20
ValidIndexArr(1) = 21
ValidIndexArr(2) = 42
ValidIndexArr(3) = 43
ValidIndexArr(4) = 64
ValidIndexArr(5) = 65
ValidIndexArr(6) = 86
ValidIndexArr(7) = 87
ValidIndexArr(8) = 108
ValidIndexArr(9) = 109
ValidIndexArr(10) = 150
ValidIndexArr(11) = 151
ValidIndexArr(12) = 152
ValidIndexArr(13) = 153
ValidIndexArr(14) = 174
ValidIndexArr(15) = 175
ValidIndexArr(16) = 196
ValidIndexArr(17) = 197
ValidIndexArr(18) = 218
ValidIndexArr(19) = 219
ValidIndexArr(20) = 240
ValidIndexArr(21) = 241
ValidIndexArr(22) = 262
ValidIndexArr(23) = 263
ValidIndexArr(24) = 284
ValidIndexArr(25) = 285
ValidIndexArr(26) = 306
ValidIndexArr(27) = 307
ValidIndexArr(28) = 328
ValidIndexArr(29) = 329
ValidIndexArr(30) = 350
ValidIndexArr(31) = 351
ValidIndexArr(32) = 372
ValidIndexArr(33) = 373
ValidIndexArr(34) = 394
ValidIndexArr(35) = 395
ValidIndexArr(36) = 416
ValidIndexArr(37) = 417
ValidIndexArr(38) = 438
ValidIndexArr(39) = 439
ValidIndexArr(40) = 460
ValidIndexArr(41) = 461
ValidIndexArr(42) = 482
ValidIndexArr(43) = 483
ValidIndexArr(44) = 504
ValidIndexArr(45) = 505
ValidIndexArr(46) = 526
ValidIndexArr(47) = 527
ValidIndexArr(48) = 548
ValidIndexArr(49) = 549

RegNameArr(0) = "AO_U1_HPM_PC_ORG"
RegNameArr(1) = "AO_U2_HPM_PC_ORG"
RegNameArr(2) = "FCM_U1_HPM_PC_ORG"
RegNameArr(3) = "FCM_U2_HPM_PC_ORG"
RegNameArr(4) = "G3D_U1_HPM_PC_ORG"
RegNameArr(5) = "G3D_U2_HPM_PC_ORG"
RegNameArr(6) = "MODEM5G_U1_HPM_PC_ORG"
RegNameArr(7) = "MODEM5G_U2_HPM_PC_ORG"
RegNameArr(8) = "MODEM_U1_HPM_PC_ORG"
RegNameArr(9) = "MODEM_U2_HPM_PC_ORG"
RegNameArr(10) = "NPU_NPU_U1_HPM_PC_ORG"
RegNameArr(11) = "NPU_WRAP_U1_HPM_PC_ORG"
RegNameArr(12) = "NPU_NPU_U2_HPM_PC_ORG"
RegNameArr(13) = "NPU_WRAP_U2_HPM_PC_ORG"
RegNameArr(14) = "PERI_U1_HPM_PC_ORG"
RegNameArr(15) = "PERI_U2_HPM_PC_ORG"
RegNameArr(16) = "ANANKE_U1_HPM_PC_ORG"
RegNameArr(17) = "ANANKE_U2_HPM_PC_ORG"
RegNameArr(18) = "ANANKE_U1_HPM_1_PC_ORG"
RegNameArr(19) = "ANANKE_U2_HPM_1_PC_ORG"
RegNameArr(20) = "ANANKE_U1_HPM_2_PC_ORG"
RegNameArr(21) = "ANANKE_U2_HPM_2_PC_ORG"
RegNameArr(22) = "ANANKE_U1_HPM_3_PC_ORG"
RegNameArr(23) = "ANANKE_U2_HPM_3_PC_ORG"
RegNameArr(24) = "BBP5G_U1_HPM_PC_ORG"
RegNameArr(25) = "BBP5G_U2_HPM_PC_ORG"
RegNameArr(26) = "ENYO_U1_HPM_PC_ORG"
RegNameArr(27) = "ENYO_U2_HPM_PC_ORG"
RegNameArr(28) = "ENYO_MIDCORE_U1_HPM_PC_ORG"
RegNameArr(29) = "ENYO_MIDCORE_U2_HPM_PC_ORG"
RegNameArr(30) = "ENYO_MIDCORE_U1_HPM_1_PC_ORG"
RegNameArr(31) = "ENYO_MIDCORE_U2_HPM_1_PC_ORG"
RegNameArr(32) = "ENYO_MIDCORE_U1_HPM_2_PC_ORG"
RegNameArr(33) = "ENYO_MIDCORE_U2_HPM_2_PC_ORG"
RegNameArr(34) = "TRYM_U1_HPM_PC_ORG"
RegNameArr(35) = "TRYM_U2_HPM_PC_ORG"
RegNameArr(36) = "TRYM_U1_HPM_1_PC_ORG"
RegNameArr(37) = "TRYM_U2_HPM_1_PC_ORG"
RegNameArr(38) = "TRYM_U1_HPM_2_PC_ORG"
RegNameArr(39) = "TRYM_U2_HPM_2_PC_ORG"
RegNameArr(40) = "TRYM_U1_HPM_3_PC_ORG"
RegNameArr(41) = "TRYM_U2_HPM_3_PC_ORG"
RegNameArr(42) = "TRYM_U1_HPM_4_PC_ORG"
RegNameArr(43) = "TRYM_U2_HPM_4_PC_ORG"
RegNameArr(44) = "TRYM_U1_HPM_5_PC_ORG"
RegNameArr(45) = "TRYM_U2_HPM_5_PC_ORG"
RegNameArr(46) = "TRYM_U1_HPM_6_PC_ORG"
RegNameArr(47) = "TRYM_U2_HPM_6_PC_ORG"
RegNameArr(48) = "TRYM_U1_HPM_7_PC_ORG"
RegNameArr(49) = "TRYM_U2_HPM_7_PC_ORG"


RegIndexArr(0) = 0
RegIndexArr(1) = 10
RegIndexArr(2) = 22
RegIndexArr(3) = 32
RegIndexArr(4) = 44
RegIndexArr(5) = 54
RegIndexArr(6) = 66
RegIndexArr(7) = 76
RegIndexArr(8) = 88
RegIndexArr(9) = 98
RegIndexArr(10) = 110
RegIndexArr(11) = 120
RegIndexArr(12) = 130
RegIndexArr(13) = 140
RegIndexArr(14) = 154
RegIndexArr(15) = 164
RegIndexArr(16) = 176
RegIndexArr(17) = 186
RegIndexArr(18) = 198
RegIndexArr(19) = 208
RegIndexArr(20) = 220
RegIndexArr(21) = 230
RegIndexArr(22) = 242
RegIndexArr(23) = 252
RegIndexArr(24) = 264
RegIndexArr(25) = 274
RegIndexArr(26) = 286
RegIndexArr(27) = 296
RegIndexArr(28) = 308
RegIndexArr(29) = 318
RegIndexArr(30) = 330
RegIndexArr(31) = 340
RegIndexArr(32) = 352
RegIndexArr(33) = 362
RegIndexArr(34) = 374
RegIndexArr(35) = 384
RegIndexArr(36) = 396
RegIndexArr(37) = 406
RegIndexArr(38) = 418
RegIndexArr(39) = 428
RegIndexArr(40) = 440
RegIndexArr(41) = 450
RegIndexArr(42) = 462
RegIndexArr(43) = 472
RegIndexArr(44) = 484
RegIndexArr(45) = 494
RegIndexArr(46) = 506
RegIndexArr(47) = 516
RegIndexArr(48) = 528
RegIndexArr(49) = 538

ValidIndex_dsp.Data = ValidIndexArr
RegIndex_dsp.Data = RegIndexArr

End Sub
