Attribute VB_Name = "VBT_A_PCIE"
Option Explicit

Function PCIEIO_FreqCntr_vbt( _
         in_RelayMode_tl As tlRelayMode, _
         in_ConfigPatName_PT As Pattern, _
         in_CapturePatName_PT As Pattern, _
         in_PLLPins_PL As PinList, _
         in_VthValue_dbl As Double, _
         in_FreqMeasWin_dbl As Double, _
         Optional DriveLoPins As PinList, _
         Optional DriveHiPins As PinList, _
         Optional DriveZPins As PinList) As Long
'this function correspond to "UFS_PLL_DCO" in the 6280

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__PCIE", "PCIEIO_FreqCntr_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i_RunPrecondition_bool                      As Boolean
    Dim t_FFSLists_str()                            As String
    Dim t_FFSListCount_lng                          As Long
    Dim t_Freq_PLD                                  As New PinListData
    Dim t_FuncResult_SBOOL                          As New SiteBoolean
    Dim i                                           As Long
    

    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl, DriveHiPins, DriveLoPins, DriveZPins

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If


    '   II. setup Voh for FreqCntr: move VOH away from VT
'    With TheHdw.Digital.Pins(in_PLLPins_PL).Levels
'        If .Value(chVol) < in_VthValue_dbl Then
'            .Value(chVoh) = in_VthValue_dbl
'        Else
'            .Value(chVol) = in_VthValue_dbl
'            .Value(chVoh) = in_VthValue_dbl
'        End If
'    End With


    '   III. Measure Freq
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If

    TheHdw.Patterns(in_CapturePatName_PT).Start
    TheHdw.Digital.Patgen.FlagWait cpuA, 0

    Call MeasureFrequency(in_PLLPins_PL, in_FreqMeasWin_dbl, t_Freq_PLD, 0)

    TheHdw.Digital.Patgen.Continue 0, cpuA
    TheHdw.Digital.Patgen.HaltWait


    '   IV. restore Voh settings, not sure if needed...
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl
    
    Call TheExec.DataManager.DecomposePinList(in_PLLPins_PL, t_FFSLists_str, t_FFSListCount_lng)

    '   V. Datalog
    ' log frequency
    
    t_FuncResult_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    TheExec.Flow.TestLimitIndex = 0
    TheExec.Flow.TestLimit ResultVal:=t_FuncResult_SBOOL, PinName:=in_CapturePatName_PT.Value, forceResults:=tlForceFlow
    

    For i = 0 To t_FFSListCount_lng - 1
        TheExec.Flow.TestLimit _
                ResultVal:=t_Freq_PLD.Pins(i), _
                forceResults:=tlForceFlow
    Next i

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function




