Attribute VB_Name = "VBT_A_DDR"
Option Explicit

Public Function DDR_SPO_ReadCode_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_DataRate_dbl As Double, _
       in_DelayMeas_dbl As Double) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__DDR", "DDR_SPO_ReadCode_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_CaptureSampleSize_lng As Long = 500
    Const i_SignalName_str As String = "DDR_SPO_ReadCode"

    Dim i_RunPrecondition_bool As Boolean

    Dim i_TestCount_lng As Long
    Dim i_RegName_str() As String

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_RdsqCycN_PLD As New PinListData
    Dim t_SPO_PLD As New PinListData
    Dim t_Jitter_PLD As New PinListData
    Dim t_Jitter_REF_PLD As New PinListData
    Dim t_Jitter_FB_PLD As New PinListData


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If


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
    
    ' retrieve pattern burst result
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite


    '   IV. Backgroud Calculation
    rundsp.D_03_CalculateSPO i_CapCodes_PLD, in_DataRate_dbl, in_DelayMeas_dbl, t_RdsqCycN_PLD, t_SPO_PLD, t_Jitter_PLD, t_Jitter_REF_PLD, t_Jitter_FB_PLD


    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName   ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit _
            ResultVal:=t_FuncResult_SBOOL, _
            TName:="", _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_RdsqCycN_PLD.Pins(i_CapturePins_PL), _
            TName:="RDQSCYC_N", _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_SPO_PLD.Pins(i_CapturePins_PL), _
            TName:="SPO", _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_Jitter_PLD.Pins(i_CapturePins_PL), _
            TName:="JITTER", _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_Jitter_REF_PLD.Pins(i_CapturePins_PL), _
            TName:="JITTER_REF", _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_Jitter_FB_PLD.Pins(i_CapturePins_PL), _
            TName:="JITTER_FB", _
            forceResults:=tlForceFlow
    

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    '   X.


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function DDR_Loopback_ReadCode_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__DDR", "DDR_Loopback_ReadCode_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_CaptureSampleSize_lng As Long = 2
    Const i_SignalName_str As String = "DDR_Loopback_ReadCode"

    Dim i_RunPrecondition_bool As Boolean

    Dim i_TestCount_lng As Long
    Dim i_RegName_str() As String

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_ZCalResult_PLD As New PinListData



    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If


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

    TheHdw.Patterns(in_CapturePatName_PT).Start
    TheHdw.Digital.Patgen.HaltWait
    ' retrieve pattern burst result
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite


    '   IV. Backgroud Calculation
    rundsp.D_07_CalDDRLPBK i_CapCodes_PLD, t_ZCalResult_PLD


    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName   ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit _
            ResultVal:=t_FuncResult_SBOOL, _
            forceResults:=tlForceFlow
    TheExec.Flow.TestLimit _
            ResultVal:=t_ZCalResult_PLD.Pins(i_CapturePins_PL), _
            TName:="zcal_result", _
            forceResults:=tlForceFlow

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    '   X.


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
