Attribute VB_Name = "VBT_A_Functional"
Option Explicit


Private Const m_CAPTUREWAIT_dbl As Double = 0.001    ' new spec settling time in sec,

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
Public Function A_Functional_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_TestPatName_PT As Pattern, _
       Optional in_CheckPatPF_str As PFType = 0, _
       Optional in_ConnectAllPins_bool As Boolean = True, _
       Optional in_LoadLevels_bool As Boolean = True, _
       Optional in_LoadTiming_bool As Boolean = True, _
       Optional in_InitWaitTime_dbl As Double = 0, _
       Optional DriveLoPins As PinList, _
           Optional DriveHiPins As PinList, _
       Optional DriveZPins As PinList, _
           Optional PrePatF As InterposeName, _
       Optional PostPatF As InterposeName) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__Basic", "A_Functional_vbt", TheExec.DataManager.InstanceName)


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming in_ConnectAllPins_bool, in_LoadLevels_bool, in_LoadTiming_bool, in_RelayMode_tl, _
                                    DriveHiPins, DriveLoPins, DriveZPins
    ' Apply AVS to Power
    Call AVS_ApplyLevel(0.0001, False)
    
    ' Register IP
    If Len(PrePatF.Value) > 0 Then Call tl_SetInterpose(TL_C_PREPATF, PrePatF.Value, "")
    If Len(PostPatF.Value) > 0 Then Call tl_SetInterpose(TL_C_POSTPATF, PostPatF.Value, "")
    ' Print the number of pattern cycles
    If glb_PrintCycleCount_bool Then Call tl_SetInterpose(TL_C_POSTPATF, "PrintCycleCount", "")
        ' Collect CheckInfo log
    If PrintInfo Then Call tl_SetInterpose(TL_C_POSTPATF, "CheckInfoCollect", "")
    
    If in_InitWaitTime_dbl > 0 Then TheHdw.Wait in_InitWaitTime_dbl


    ''    'Declaration
    ''    Dim i_pfType As PFType
    ''
    ''    If in_CheckPatPF_str = "pfAlways" Then
    ''        i_pfType = pfAlways
    ''    ElseIf in_CheckPatPF_str = "pfNever" Then
    ''        i_pfType = pfNever
    ''    ElseIf in_CheckPatPF_str = "pfFailsOnly" Then
    ''        i_pfType = pfFailsOnly
    ''    Else
    ''        i_pfType = pfAlways    ' default value
    ''    End If


    '   III. Begin the test
    TheHdw.Patterns(in_TestPatName_PT).test in_CheckPatPF_str, 0


    '   X.
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF, TL_C_POSTPATBPF)

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function A_Functional_Por_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_TestPatName_PT As Pattern, _
       Optional in_CheckPatPF_str As PFType = 0, _
       Optional in_ConnectAllPins_bool As Boolean = True, _
       Optional in_LoadLevels_bool As Boolean = True, _
       Optional in_LoadTiming_bool As Boolean = True, _
       Optional in_InitWaitTime_dbl As Double = 0, _
       Optional DriveLoPins As PinList, _
           Optional DriveHiPins As PinList, _
       Optional DriveZPins As PinList, _
           Optional PrePatF As InterposeName, _
       Optional PostPatF As InterposeName) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__Basic", "A_Functional_vbt", TheExec.DataManager.InstanceName)


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming in_ConnectAllPins_bool, in_LoadLevels_bool, in_LoadTiming_bool, in_RelayMode_tl, _
                                    DriveHiPins, DriveLoPins, DriveZPins
    
    
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = True
'    TheHdw.Protocol.Ports("CLK_38M4").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_38M4").IdleWait
'
    With TheHdw.Digital.Pins("CLK_32K").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001
    With TheHdw.Digital.Pins("CLK_38M4").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    TheHdw.Wait 0.001
    
    ' Register IP
    If Len(PrePatF.Value) > 0 Then Call tl_SetInterpose(TL_C_PREPATF, PrePatF.Value, "")
    If Len(PostPatF.Value) > 0 Then Call tl_SetInterpose(TL_C_POSTPATF, PostPatF.Value, "")
    ' Print the number of pattern cycles
    If glb_PrintCycleCount_bool Then Call tl_SetInterpose(TL_C_POSTPATF, "PrintCycleCount", "")
        ' Collect CheckInfo log
    If PrintInfo Then Call tl_SetInterpose(TL_C_POSTPATF, "CheckInfoCollect", "")
    
    If in_InitWaitTime_dbl > 0 Then TheHdw.Wait in_InitWaitTime_dbl


    ''    'Declaration
    ''    Dim i_pfType As PFType
    ''
    ''    If in_CheckPatPF_str = "pfAlways" Then
    ''        i_pfType = pfAlways
    ''    ElseIf in_CheckPatPF_str = "pfNever" Then
    ''        i_pfType = pfNever
    ''    ElseIf in_CheckPatPF_str = "pfFailsOnly" Then
    ''        i_pfType = pfFailsOnly
    ''    Else
    ''        i_pfType = pfAlways    ' default value
    ''    End If


    '   III. Begin the test
    TheHdw.Patterns(in_TestPatName_PT).test in_CheckPatPF_str, 0


    '   X.
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF, TL_C_POSTPATBPF)
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = False
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = False
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    TheHdw.Digital.Pins("CLK_38M4").FreeRunningClock.Stop
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


