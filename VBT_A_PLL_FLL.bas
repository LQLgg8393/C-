Attribute VB_Name = "VBT_A_PLL_FLL"
Option Explicit
Global tmpHashTable0 As New clsHashTable
Global g_RegName_str() As String
Global g_FreqPin_str() As String

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'


Public Function PLL_ReadUnlockCounter_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_CapturePatName_PT As Pattern, _
       Optional in_ConfigPatName_PT As Pattern) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__PLL_FLL", "PLL_ReadUnlockCounter_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long
    Dim Site As Variant
    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_CaptureSampleSize_lng As Long = 192   '6*32bit
    Const i_SignalName_str As String = "PLL_ReadUnlockCounter"

    Dim i_RunPrecondition_bool As Boolean

    Dim i_TestCount_lng As Long
    Dim i_RegName_str() As String
    Dim i_RegisterNames_str As String

    Dim i_RegData_PLD As New PinListData
    Dim t_RegData_DSP As New DSPWave


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl
    
    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001
    
    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    i_RegisterNames_str = "1ST_READ_0XFFF0592C,1ST_READ_0XFFF05930,1ST_READ_0XFFF05934,2ND_READ_0XFFF0592C,2ND_READ_0XFFF05930,2ND_READ_0XFFF05934"
    i_RegName_str = Split(i_RegisterNames_str, ",")
    i_TestCount_lng = UBound(i_RegName_str) + 1

    '   II. Setup DSSC
    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         i_CaptureSampleSize_lng, _
         i_CapCodes_PLD)


    '   III. Start the Pattern(s)
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If
    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(in_CapturePatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If
    '   IV. Foregroud Calculation
    ' 8 x 16bit word, should be modified to parallel capture so DSP is not needed!
    Call rundsp.D_06_Cal_PLL_Unlock(i_CapCodes_PLD, i_RegData_PLD)
    
    t_RegData_DSP = i_RegData_PLD.Pins(i_CapturePins_PL)

    '   V. Datalog
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    For i = 0 To i_TestCount_lng - 1
        '        TheExec.Flow.TestLimitIndex = 0
        TheExec.Flow.TestLimit _
                ResultVal:=t_RegData_DSP.Element(i), _
                TName:="PLL_UNLOCK_" + i_RegName_str(i), _
                PinName:=i_CapturePins_PL, _
                forceResults:=tlForceFlow
    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    '   X.
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Function PLL_FreqCntr_vbt( _
         in_RelayMode_tl As tlRelayMode, _
         in_CapturePatName_PT As Pattern, _
         in_RegisterNames_str As String, _
         in_PLLPins_PL As PinList, _
         in_CaptureSampleSize_lng As Long, _
         in_FreqMeasWin_dbl As Double, _
         Optional in_ConfigPatName_PT As Pattern, _
         Optional in_VthValue_dbl As Double) As Long
' almost identical with SVFD_FFC_FreqCntr_vbt, added support for Voh change

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__PLL_FLL", "PLL_FreqCntr_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i                       As Long
    Dim Site                    As Variant

    Dim t_CapCodes_PLD          As New PinListData
    Dim i_CapturePins_PL        As New PinList

    Const i_SignalName_str      As String = "PLL_FreqCntr"

    Dim i_RunPrecondition_bool  As Boolean

    Dim i_TestCount_lng         As Long
    Dim i_MeasurePins_str()     As String
    Dim t_RegData_DSP           As New DSPWave
    Dim i_RegName_str()         As String
    
    Dim t_PLLFreq_PLD           As New PinListData

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001
    
    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    in_RegisterNames_str = Trim(in_RegisterNames_str)
    i_RegName_str = Split(in_RegisterNames_str, ",")

    Call TheExec.DataManager.DecomposePinList(in_PLLPins_PL, i_MeasurePins_str, i_TestCount_lng)
    
    If i_TestCount_lng <> UBound(i_RegName_str) + 1 Then
        ' check the instance sheet and correct the 'PLLPins' or 'RegisterNames'
        TheExec.Datalog.WriteComment _
                "FreqCntr Pin count is: " + CStr(i_TestCount_lng) + _
                                          ". while regsiter number is: " + CStr(UBound(i_RegName_str) + 1)
        Stop
        Exit Function
    End If


'   II. Setup DSSC  &  setup Voh for FreqCntr: move VOH away from VT
'    With TheHdw.Digital.Pins(in_PLLPins_PL).Levels
'        If .Value(chVol) < in_VthValue_dbl Then
'            .Value(chVoh) = in_VthValue_dbl
'        Else
'            .Value(chVol) = in_VthValue_dbl
'            .Value(chVoh) = in_VthValue_dbl
'        End If
'    End With

    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng, _
         t_CapCodes_PLD)


    '   III. Begin the test
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(in_CapturePatName_PT).Start
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
    End If
    
    Call MeasureFrequency(in_PLLPins_PL, in_FreqMeasWin_dbl, t_PLLFreq_PLD)

    TheHdw.Digital.Patgen.Continue 0, cpuA
    TheHdw.Digital.Patgen.HaltWait
       
'   restore Voh settings, not sure if needed...
   TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

'   IV. check if input register name is smaller than capture size
    For Each Site In TheExec.Sites
        i_TestCount_lng = t_CapCodes_PLD.Pins(i_CapturePins_PL).Value.SampleSize
        If i_TestCount_lng - 1 > UBound(i_RegName_str) Then
            ' check the instance sheet and correct the 'CaptureSampleSize' or 'RegisterNames'
            TheExec.Datalog.WriteComment _
                    "Captured sample size is: " + CStr(i_TestCount_lng) + _
                                                ". while regsiter number is: " + CStr(UBound(i_RegName_str) + 1)

            i_TestCount_lng = UBound(i_RegName_str) + 1
        End If
        t_RegData_DSP = t_CapCodes_PLD.Pins(i_CapturePins_PL).Value
    Next


    '   V. Datalog
    'log register value
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019
    
    For i = 0 To i_TestCount_lng - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_RegData_DSP.Element(i), _
                TName:="PLL_LOCK_" + i_RegName_str(i), _
                PinName:=i_CapturePins_PL, _
                forceResults:=tlForceFlow
    Next i

    ' log frequency
    For i = 0 To t_PLLFreq_PLD.Pins.count - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_PLLFreq_PLD.Pins(i_MeasurePins_str(i)), _
                PinName:=i_MeasurePins_str(i), _
                TName:="PLL_FREQ_" + i_RegName_str(i), _
                forceResults:=tlForceFlow
    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Function PLL_FreqCntr_CMEM_vbt( _
         in_RelayMode_tl As tlRelayMode, _
         in_CapturePatName_PT As Pattern, _
         in_PLLLowFREQPins_PL As PinList, _
         in_PLLNonLowPins_PL As PinList, _
         in_CaptureSampleSize_LowFreq_lng As Long, _
         in_CaptureSampleSize_LockPin_lng As Long, _
         in_RegisterNames_str As String, _
         in_FreqMeasWin_dbl As Double, _
         Optional in_ConfigPatName_PT As Pattern, _
         Optional in_VthValue_dbl As Double, _
        Optional HizPins As PinList) As Long
' almost identical with SVFD_FFC_FreqCntr_vbt, added support for Voh change
    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__PLL_FLL", "PLL_FreqCntr_vbt", TheExec.DataManager.InstanceName)
    '   Declarations
    Dim i                               As Long
    Dim Site                            As Variant

    Dim t_CapCodes_PLD                  As New PinListData
    Dim i_CapturePins_PL                As New PinList
    Dim i_CapCmem_PLD                   As New PinListData
    Dim t_CapFreq_PLD                   As New PinListData

    Const i_SignalName_LockPin_str      As String = "PLL_LockPin_FreqCntr"
    Dim i_RunPrecondition_bool          As Boolean

    Dim i_TestCount_LowFreq_lng         As Long
    Dim i_TestCount_NonLowPin_lng       As Long
    Dim i_MeasurePins_LowFreq_str()     As String
    Dim i_MeasurePins_NonLowPin_str()   As String
    Dim i_RegName_str()                 As String
    Dim t_RegData_DSP                   As New DSPWave
    Dim t_FreqData_DSP                  As New DSPWave
    
    Dim i_tmp_DSP                       As New DSPWave
    Dim i_FreqPin_str                   As Variant
    Dim t_PLLFreq_PLD                   As New PinListData
    Dim t_PLLFreqL_PLD                  As New PinListData
    Dim t_PLLLock_PLD                   As New PinListData
    Dim CMEM_CapturedCycles             As Long
    in_RegisterNames_str = Trim(in_RegisterNames_str)
    i_RegName_str = Split(in_RegisterNames_str, ",")
    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl, , , HizPins

    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001
    
    'init PLL RegName and FREQ Pin
    If tmpHashTable0.count > 0 Then
    Else: Call Init_PLLLockPin_Array
    End If
    
    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    Call TheExec.DataManager.DecomposePinList(in_PLLLowFREQPins_PL, i_MeasurePins_LowFreq_str, i_TestCount_LowFreq_lng)
    Call TheExec.DataManager.DecomposePinList(in_PLLNonLowPins_PL, i_MeasurePins_NonLowPin_str, i_TestCount_NonLowPin_lng)
    If in_CaptureSampleSize_LockPin_lng <> UBound(i_RegName_str) + 1 Then
        ' check the instance sheet and correct the 'PLLPins' or 'RegisterNames'
        TheExec.Datalog.WriteComment _
                "Lock Pin count is: " + CStr(in_CaptureSampleSize_LockPin_lng) + _
                                          ". while regsiter number is: " + CStr(UBound(i_RegName_str) + 1)
        Stop
        Exit Function
    End If
    
    ' init the DSPWave
    t_RegData_DSP.CreateConstant 0, in_CaptureSampleSize_LockPin_lng, DspLong
    t_FreqData_DSP.CreateConstant 0, in_CaptureSampleSize_LowFreq_lng, DspLong
    i_CapCmem_PLD.AddPin ("JTAG_TDO")
    i_CapCmem_PLD.Pins("JTAG_TDO").Value = t_RegData_DSP
    For Each i_FreqPin_str In i_MeasurePins_LowFreq_str
        t_CapFreq_PLD.AddPin (i_FreqPin_str)
        t_CapFreq_PLD.Pins(i_FreqPin_str).Value = t_FreqData_DSP
    Next
    
    'cmem setup
    'TheHdw.Digital.CMEM.CentralFields = tlCMEMAbsoluteCycle
    Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone)
    Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptSTV, tlCMEMCaptureSourceDutData)

    '   III. Begin the test
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(in_CapturePatName_PT).Start
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
    End If
    
    Call MeasureFrequency(in_PLLNonLowPins_PL, in_FreqMeasWin_dbl, t_PLLFreq_PLD)

    TheHdw.Digital.Patgen.Continue 0, cpuA
    TheHdw.Digital.Patgen.HaltWait
    If TheExec.TesterMode = testModeOnline Then
        CMEM_CapturedCycles = TheHdw.Digital.CMEM.CapturedCycles
        If CMEM_CapturedCycles = in_CaptureSampleSize_LockPin_lng + in_CaptureSampleSize_LowFreq_lng Then
            i_CapCmem_PLD = TheHdw.Digital.Pins("JTAG_TDO").CMEM.Data(0, in_CaptureSampleSize_LockPin_lng, tlCMEMNoPackData)
            t_CapFreq_PLD = TheHdw.Digital.Pins(in_PLLLowFREQPins_PL).CMEM.Data(in_CaptureSampleSize_LockPin_lng, in_CaptureSampleSize_LowFreq_lng, tlCMEMPackData)
        Else
            TheExec.Datalog.WriteComment _
            "CMEM CapturedCycles : " + CStr(CMEM_CapturedCycles) + _
                                ". Total CapturedCycles is: " + CStr(in_CaptureSampleSize_LockPin_lng + in_CaptureSampleSize_LowFreq_lng)
            Exit Function
        End If
    End If
'   IV.  Calculation
    Dim i_tmp_lng() As Long
    Dim i_tmp_byt() As Byte

    For Each Site In TheExec.Sites.Active
        If TheExec.TesterMode = testModeOnline Then
            i_tmp_byt = i_CapCmem_PLD.Pins("JTAG_TDO").Value
            ReDim i_tmp_lng(UBound(i_tmp_byt))
            For i = 0 To UBound(i_tmp_byt)
                i_tmp_lng(i) = i_tmp_byt(i)
            Next i
            t_RegData_DSP.Data = i_tmp_lng
        End If
        For i = 0 To i_TestCount_LowFreq_lng - 1
            If TheExec.TesterMode = testModeOnline Then
                i_tmp_DSP.Data = t_CapFreq_PLD.Pins(i_MeasurePins_LowFreq_str(i)).Value
            ElseIf TheExec.TesterMode = testModeOffline Then
                i_tmp_DSP.CreateConstant 0, in_CaptureSampleSize_LockPin_lng, DspLong
            End If
                t_CapFreq_PLD.Pins(i_MeasurePins_LowFreq_str(i)).Value = i_tmp_DSP
        Next i
    Next Site
    For i = 0 To i_TestCount_LowFreq_lng - 1
        t_PLLFreqL_PLD.AddPin (t_CapFreq_PLD.Pins(i))
    Next i
    
    Dim i_Period_dbl As Double
    i_Period_dbl = TheHdw.Digital.Timing.Period("timeplate_1_1")
    rundsp.CalcFreq t_CapFreq_PLD, i_Period_dbl, t_PLLFreqL_PLD

    '   V. Datalog
    'log register value
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019
    Dim Key As Variant

    For i = 0 To in_CaptureSampleSize_LockPin_lng - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_RegData_DSP.Element(i), _
                TName:="PLL_LOCK_" + i_RegName_str(i), _
                PinName:="JTAG_TDO", _
                forceResults:=tlForceFlow
    Next i

    For i = 0 To t_PLLFreqL_PLD.Pins.count - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_PLLFreqL_PLD.Pins(i_MeasurePins_LowFreq_str(i)), _
                PinName:=i_MeasurePins_LowFreq_str(i), _
                TName:="PLL_FREQ_" + tmpHashTable0.Value(i_MeasurePins_LowFreq_str(i)), _
                forceResults:=tlForceFlow
    Next i
    For i = 0 To t_PLLFreq_PLD.Pins.count - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_PLLFreq_PLD.Pins(i_MeasurePins_NonLowPin_str(i)), _
                PinName:=i_MeasurePins_NonLowPin_str(i), _
                TName:="PLL_FREQ_" + tmpHashTable0.Value(i_MeasurePins_NonLowPin_str(i)), _
                forceResults:=tlForceFlow
    Next i
    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone)

    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Sub Init_PLLLockPin_Array()

    ReDim g_RegName_str(24)
    ReDim g_FreqPin_str(24)
    g_RegName_str(0) = "APLL0"
    g_RegName_str(1) = "APLL1"
    g_RegName_str(2) = "APLL2"
    g_RegName_str(3) = "APLL3"
    g_RegName_str(4) = "APLL5"
    g_RegName_str(5) = "FNPLL1"
    g_RegName_str(6) = "FNPLL4"
    g_RegName_str(7) = "BINT"
    g_RegName_str(8) = "PCIE_FNPLL"
    g_RegName_str(9) = "PPLL1"
    g_RegName_str(10) = "PPLL2"
    g_RegName_str(11) = "PPLL3"
    g_RegName_str(12) = "PPLL4"
    g_RegName_str(13) = "SCPLL0"
    g_RegName_str(14) = "SCPLL1"
    g_RegName_str(15) = "SCPLL2"
    g_RegName_str(16) = "SCPLL3"
    g_RegName_str(17) = "SCPLL4"
    g_RegName_str(18) = "SCPLL5"
    g_RegName_str(19) = "SCPLL6"
    g_RegName_str(20) = "SCPLL7"
    g_RegName_str(21) = "SCPLL8"
    g_RegName_str(22) = "SCPLL9"
    g_RegName_str(23) = "SPLL"
    g_RegName_str(24) = "FLL"
    g_FreqPin_str(0) = "GPIO_BBA_6"
    g_FreqPin_str(1) = "GPIO_BBA_7"
    g_FreqPin_str(2) = "GPIO_069"
    g_FreqPin_str(3) = "GPIO_BBA_2"
    g_FreqPin_str(4) = "GPIO_BBA_3"
    g_FreqPin_str(5) = "GPIO_068"
    g_FreqPin_str(6) = "GPIO_052"
    g_FreqPin_str(7) = "UART4_RXD"
    g_FreqPin_str(8) = "RF0_ALINK_DATA"
    g_FreqPin_str(9) = "GPIO_039"
    g_FreqPin_str(10) = "GPIO_038"
    g_FreqPin_str(11) = "GPIO_048"
    g_FreqPin_str(12) = "GPIO_049"
    g_FreqPin_str(13) = "LTE_GPS_TIM_IND"
    g_FreqPin_str(14) = "RF0_MCSYNC"
    g_FreqPin_str(15) = "RF0_WDOG_INT"
    g_FreqPin_str(16) = "RF0_RESET_N"
    g_FreqPin_str(17) = "UART6_RXD"
    g_FreqPin_str(18) = "UART6_TXD"
    g_FreqPin_str(19) = "UART6_CTS_N"
    g_FreqPin_str(20) = "UART6_RTS_N"
    g_FreqPin_str(21) = "GPIO_230"
    g_FreqPin_str(22) = "GPIO_229"
    g_FreqPin_str(23) = "UART4_RTS_N"
    g_FreqPin_str(24) = "UART4_CTS_N"
    
    Dim i As Long
    For i = 0 To UBound(g_FreqPin_str)
        If tmpHashTable0.KeyExists(g_FreqPin_str(i)) Then
        Else
            Call tmpHashTable0.Add(g_FreqPin_str(i), g_RegName_str(i))
        End If
    Next i

End Sub
