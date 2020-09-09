Attribute VB_Name = "VBT_A_USB"
Option Explicit
Private Const m_DP_Pin As String = "USB2_DP"
Private Const m_DM_Pin As String = "USB2_DM"


Public Function USB_CharDevice_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_PatName_PT As Pattern) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__USB", "USB_CharDevice_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i_DPPins As New PinList
    Dim i_DMPins As New PinList
    Dim i_DiffPins As New PinList

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_V_DP_0A_PLD As New PinListData
    Dim t_I_DP_500mV_PLD As New PinListData
    Dim t_I_DM_150mV_PLD As New PinListData
    Dim t_I_DM_2V_PLD As New PinListData

    '   1. run pattern
    '   2. FIMV USB_DP @ 0A,       expecting between 0.5V~0.7V
    '   3. FVMI USB_DP @ 500mV,    expecting > 250uA;
    '   4. FVMI USB_DM @ 0.15V     expecting between 25uA~175uA
    '   5. FVMI USB_DM @ 2V        expecting between 25uA~175uA;

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    i_DPPins.Value = m_DP_Pin
    i_DMPins.Value = m_DM_Pin
    i_DiffPins.Value = m_DP_Pin + "," + m_DM_Pin


    '   1. run pattern
    TheHdw.Patterns(in_PatName_PT).Start
    TheHdw.Digital.Patgen.HaltWait
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    'Open PE
    TheHdw.Digital.Pins(i_DiffPins).Disconnect

    '   2. FIMV USB_DP @ 0A,       expecting between 0.5V~0.7V
    Call PPMU_FIMV(i_DPPins, 0, t_V_DP_0A_PLD, , , True, False, False, 0.00002)

    '   3. FVMI USB_DP @ 500mV,    expecting > 250uA;
    Call PPMU_FVMI(i_DPPins, 0.5, t_I_DP_500mV_PLD, , , False, True, False, 0.002)

    '   4. FVMI USB_DM @ 0.15V     expecting between 25uA~175uA
    Call PPMU_FVMI(i_DMPins, 0.15, t_I_DM_150mV_PLD, , , True, False, False, 0.002)

    '   5. FVMI USB_DM @ 2V        expecting between 25uA~175uA;
    Call PPMU_FVMI(i_DMPins, 2, t_I_DM_2V_PLD, , , False, True, False, 0.002)

    'Switch back to PE, not sure if needed...[TTR]
    TheHdw.Digital.Pins(i_DiffPins).Connect


    '   III. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit ResultVal:=t_FuncResult_SBOOL, TName:="USB2_FUNCTIONAL_RESULTS", PinName:=in_PatName_PT.Value, forceResults:=tlForceFlow

    TheExec.Flow.TestLimit ResultVal:=t_V_DP_0A_PLD, TName:="VAL_NO_LOAD", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_I_DP_500mV_PLD.Math.Negate, TName:="VAL_500MV", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_I_DM_150mV_PLD, TName:="VAL_150MV", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_I_DM_2V_PLD, TName:="VAL_2000MV", forceResults:=tlForceFlow

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function USB_DataCon_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_PatName_PT As Pattern) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__USB", "USB_DataCon_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i_DPPins As New PinList
    Dim i_DMPins As New PinList
    Dim i_DiffPins As New PinList

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_V_DP_0A_PLD As New PinListData
    Dim t_V_DP_n50uA_PLD As New PinListData
    Dim t_I_DP_0V_PLD As New PinListData
    Dim t_I_DM_50mV_PLD As New PinListData
    Dim t_I_DM_2V_PLD As New PinListData
    Dim t_R_DM_2V_50mV_PLD As New PinListData


    '   1. run pattern
    '   2. FIMV USB_DP @ 0A,        expecting > 2V
    '   3. FIMV USB_DP @ -50uA,     expecting < 0.8V  ## crossed out in TP ##
    '   4. FVMI USB_DP @ 0V,        expecting between 10uA~16uAc;
    '   5.1 FVMI USB_DM @ 50mV      expecting between 25uA~175uA
    '   5.2 FVMI USB_DM @ 2V        expecting between 25uA~175uA
    '   5.3 Calc R = DeltaV/DeltaI  expecting between 4.25K?~24.8K??


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    i_DPPins.Value = m_DP_Pin
    i_DMPins.Value = m_DM_Pin
    i_DiffPins.Value = m_DP_Pin + "," + m_DM_Pin


    '   1. run pattern
    TheHdw.Patterns(in_PatName_PT).Start
    TheHdw.Digital.Patgen.HaltWait
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    'Open PE
    TheHdw.Digital.Pins(i_DiffPins).Disconnect

    '   2. FIMV USB_DP @ 0A,        expecting > 2V
    Call PPMU_FIMV(i_DPPins, 0, t_V_DP_0A_PLD, , , True, False, False, 0.00002)


    '   3. FIMV USB_DP @ -50uA,     expecting < 0.8V  ## crossed out in TP ##
    Call PPMU_FIMV(i_DPPins, -0.00005, t_V_DP_n50uA_PLD, , , False, False, False, 0.0002)

    '   4. FVMI USB_DP @ 0V,        expecting between 10uA~16uA
    Call PPMU_FVMI(i_DPPins, 0, t_I_DP_0V_PLD, 0.002, , False, True, False, 0.0002)


    '   5.1 FVMI USB_DM @ 50mV      expecting between 25uA~175uA
    Call PPMU_FVMI(i_DMPins, 0.05, t_I_DM_50mV_PLD, , , True, False, False, 0.0002)


    '   5.2 FVMI USB_DM @ 2V        expecting between 25uA~175uA
    Call PPMU_FVMI(i_DMPins, 2, t_I_DM_2V_PLD, , , False, True, False, 0.0002)


    '   5.3 Calc R = DeltaV/DeltaI  expecting between 4.25K?~24.8K??
    rundsp.USB_Calc_R 2, 0.05, t_I_DM_2V_PLD, t_I_DM_50mV_PLD, t_R_DM_2V_50mV_PLD

    'Switch back to PE, not sure if needed...[TTR]
    TheHdw.Digital.Pins(i_DiffPins).Connect


    '   III. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit ResultVal:=t_FuncResult_SBOOL, TName:="USB2_FUNCTIONAL_RESULTS", PinName:=in_PatName_PT.Value, forceResults:=tlForceFlow

    TheExec.Flow.TestLimit ResultVal:=t_V_DP_0A_PLD, TName:="VAL_200PF", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_V_DP_n50uA_PLD, TName:="VAL_15K", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_I_DP_0V_PLD, TName:="VAL_GROUND", forceResults:=tlForceFlow
    '    THEEXEC.FLOW.TESTLIMIT RESULTVAL:=T_I_DM_50MV_PLD, FORCERESULTS:=TLFORCEFLOW
    '    THEEXEC.FLOW.TESTLIMIT RESULTVAL:=T_I_DM_2V_PLD, FORCERESULTS:=TLFORCEFLOW
    TheExec.Flow.TestLimit ResultVal:=t_R_DM_2V_50mV_PLD, TName:="VAL_2000MV", forceResults:=tlForceFlow

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function USB_Function_RxHS_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_PatName_PT As Pattern, _
       in_V_DP_dbl As Double, _
       in_V_DM_dbl As Double) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__USB", "USB_Function_RxHS_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i_DPPins As New PinList
    Dim i_DMPins As New PinList
    Dim i_DiffPins As New PinList

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_I_DP_0V_PLD As New PinListData
    Dim t_I_DM_0V_PLD As New PinListData

    '   1. Apply Vcom = 600mV(50mV), Vdiff = 150mV on "USB2_DP", "USB2_DM"
    '       USB_DP @ (Vcom+Vdiff)/2=375mV
    '       USB_DM @ (Vcom-Vdiff)/2=225mV
    '   2. run pattern

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    i_DPPins.Value = m_DP_Pin
    i_DMPins.Value = m_DM_Pin
    i_DiffPins.Value = m_DP_Pin + "," + m_DM_Pin


    '   1. Apply Vcom = 600mV(50mV), Vdiff = 150mV on "USB2_DP", "USB2_DM"
    '       USB_DP @ (Vcom+Vdiff)/2=375mV
    '       USB_DM @ (Vcom-Vdiff)/2=225mV
    Call PPMU_FVMI(i_DPPins, in_V_DP_dbl, t_I_DP_0V_PLD, , , True, False, True)
    Call PPMU_FVMI(i_DMPins, in_V_DM_dbl, t_I_DM_0V_PLD, , , True, False, True)


    '   2. run pattern
    TheHdw.Patterns(in_PatName_PT).Start
    TheHdw.Digital.Patgen.HaltWait
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite

    '   3.Switch back to PE, not sure if needed...[TTR]
    TheHdw.PPMU(i_DiffPins).Disconnect
    TheHdw.Digital.Pins(i_DiffPins).Connect


    '   III. Datalog
    TheExec.Flow.TestLimit ResultVal:=t_FuncResult_SBOOL, TName:="USB2_FUNCTIONAL_RESULTS", PinName:=in_PatName_PT.Value, forceResults:=tlForceFlow

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function USB_DC_TxHS_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_PatName_PT As Pattern) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__USB", "USB_DC_TxHS_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i_DPPins As New PinList
    Dim i_DMPins As New PinList
    Dim i_DiffPins As New PinList

    Dim t_FuncResult_SBOOL As New SiteBoolean
    Dim t_V_DP_0A_PLD As New PinListData
    Dim t_V_DM_0A_PLD As New PinListData
    Dim t_V_DPDM_0A_PLD As New PinListData


    '   1. run pattern
    '   2. FIMV USB_DP @ 0A,        expecting ...
    '   3. FIMV USB_DM @ 0A,        expecting ...
    
    'connect PE nad PPMU at the same time
    TheHdw.PPMU.AllowPPMUFuncRelayConnection True, False

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    i_DPPins.Value = m_DP_Pin
    i_DMPins.Value = m_DM_Pin
    i_DiffPins.Value = m_DP_Pin + "," + m_DM_Pin


    '   1. run pattern
    TheHdw.Patterns(in_PatName_PT).Start
    TheHdw.Digital.Patgen.HaltWait
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite

    '   2. FIMV USB_DP @ 0A,        expecting ...
    '   3. FIMV USB_DM @ 0A,        expecting ...
    Call PPMU_FIMV(i_DiffPins, 0, t_V_DPDM_0A_PLD, , , True, True, False, 0.00002)


    '   III. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    '''    TheExec.Flow.TestLimit ResultVal:=t_V_DPDM_0A_PLD, forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_V_DPDM_0A_PLD.Pins(m_DP_Pin), TName:="USB2_DP", forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=t_V_DPDM_0A_PLD.Pins(m_DM_Pin), TName:="USB2_DM", forceResults:=tlForceFlow
    'TheExec.Flow.TestLimitIndex = 2
    TheExec.Flow.TestLimit ResultVal:=t_FuncResult_SBOOL, TName:="USB2_FUNCTIONAL_RESULTS", PinName:=in_PatName_PT.Value, forceResults:=tlForceFlow

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    
    TheHdw.PPMU.AllowPPMUFuncRelayConnection False, False

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

