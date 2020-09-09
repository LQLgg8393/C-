Attribute VB_Name = "VBT_A_TSensor"
Option Explicit

Global setTemp As Double
Global currTemp As Double
Global Global_Djtag_Mean_DSP As New DSPWave


Public Function TSensor_DJTAG_DSSC_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_CaptureSampleSize_lng As Long, _
       in_CaptureModuleName_str As String, _
       Optional FlagEfuseUpdate As Boolean = False) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__TSensor", "TSensor_DJTAG_DSSC_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long

    Dim Site As Variant

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_SignalName_str As String = "TSENSOR_DJTAG_ReadCode"

    Dim i_RunPrecondition_bool As Boolean

    Dim i_ModuleName_str() As String
    Dim i_TestName_str() As String
    Dim i_TestCount_lng As Long

    Dim i_Chuck_Temp_SDBL As New SiteDouble


    Dim i_Temp_Mean_PLD As New PinListData
    Dim t_Temp_Mean_DSP As New DSPWave
    Dim i_Temp_Delta_Mean_PLD As New PinListData
    Dim t_Temp_Delta_Mean_DSP As New DSPWave

    Dim i_EfuseItem_str() As String
    Dim i_EfuseAddr_lng() As Long


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

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
    
    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    Call GetDJTAGAddress(in_CaptureModuleName_str, i_ModuleName_str, i_TestName_str, i_EfuseItem_str, i_EfuseAddr_lng)
    i_TestCount_lng = UBound(i_ModuleName_str)
    If i_TestCount_lng + 1 <> in_CaptureSampleSize_lng Then
        MsgBox "CaptureSampleSize should equal to number of Modules, check Instance!"
    End If

    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites
            i_Chuck_Temp_SDBL = GetChuckTemp_dbl
        Next
    Else
        i_Chuck_Temp_SDBL = GlobalVariable_Chuck_Temp
    End If

    '   II. Setup DSSC
    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng, _
         i_CapCodes_PLD)


    '   III. Begin the test
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If

    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(in_CapturePatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If

    '   IV. Backgroud Calculation
    rundsp.E_06_TSensor_DJTAG_Cal _
            i_CapCodes_PLD, _
            i_Chuck_Temp_SDBL, _
            i_Temp_Mean_PLD, _
            i_Temp_Delta_Mean_PLD
    'every DSP has a potential risk of incomplete samples...
    For Each Site In TheExec.Sites
        t_Temp_Mean_DSP = i_Temp_Mean_PLD.Pins(i_CapturePins_PL).Value
        t_Temp_Delta_Mean_DSP = i_Temp_Delta_Mean_PLD.Pins(i_CapturePins_PL).Value
    Next Site

    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    For i = 0 To i_TestCount_lng

        TheExec.Flow.TestLimit _
                ResultVal:=t_Temp_Mean_DSP.Element(i), _
                TName:="TS_MEAN_" + i_ModuleName_str(i), _
                PinName:=i_CapturePins_PL, _
                forceResults:=tlForceFlow

        TheExec.Flow.TestLimit _
                ResultVal:=t_Temp_Delta_Mean_DSP.Element(i), _
                TName:="TS_DELTA_" + i_ModuleName_str(i), _
                PinName:=i_CapturePins_PL, _
                forceResults:=tlForceFlow

    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
    '   X.

    'for glitch test use
    Global_Djtag_Mean_DSP = t_Temp_Mean_DSP

    'Update EFUSE flag
    For Each Site In TheExec.Sites
        For i = 0 To i_TestCount_lng
            If FlagEfuseUpdate = True Then
                Call hiefuse.MULT_EFUSE_ITEM_UPDATE(i_EfuseItem_str(i), 0, Site, t_Temp_Delta_Mean_DSP.Element(i), 0)
                Select Case i_ModuleName_str(i)
                    Case "LCL0_FCM0"
                        gTsrFCM0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_CPUB0"
                        gTsrRMTCPUB0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_CPUM0"
                        gTsrRMTCPUM0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_CPUM1"
                        gTsrRMTCPUM1Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_CPUM2"
                        gTsrRMTCPUM2Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_FCM0"
                        gTsrRMTFCM0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_FCM1"
                        gTsrRMTFCM1Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL0_RMT_DDRA"
                        gTsrRMTDDRAValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_MDM0"
                        gTsrMDM0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_RMT_NCSI"
                        gTsrRMTNCSIValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_RMT_NCHE"
                        gTsrRMTNCHEValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_RMT_NPDT"
                        gTsrRMTNPDTValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_RMT_5GTOP"
                        gTsrRMT5GTOPValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL1_RMT_4GTOP"
                        gTsrRMT4GTOPValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_G3D"
                        gTsrG3DValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_GPU0"
                        gTsrRMTGPU0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_GPU1"
                        gTsrRMTGPU1Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_NPU0"
                        gTsrRMTNPU0Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_ISP"
                        gTsrRMTISPValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_M1"
                        gTsrRMTM1Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_M2"
                        gTsrRMTM2Value = t_Temp_Delta_Mean_DSP.Element(i)

                    Case "LCL2_RMT_DDRB"
                        gTsrRMTDDRBValue = t_Temp_Delta_Mean_DSP.Element(i)

                    Case Else
                        If TheExec.RunMode = runModeProduction Then
                            TheExec.AddOutput "TSENSOR_DJTAG ERROR: Unknown Module Name"
                            GoTo errHandler
                        Else
                            MsgBox "TSENSOR_DJTAG ERROR: Unknown Module Name"
                            Stop
                        End If
                End Select
            End If
        Next i
    Next Site
    
    TheHdw.Digital.Pins("CLK_32K").FreeRunningClock.Stop
    TheHdw.Digital.Pins("CLK_38M4").FreeRunningClock.Stop
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function GetDJTAGAddress( _
        in_ModuleName_str As String, _
        ByRef out_ModuleName_str() As String, _
        ByRef out_TestName_str() As String, _
        ByRef out_EfuseItem_str() As String, _
        ByRef out_DJTAGADDR_lng() As Long _
      ) As Long
' this function need to be updated...
    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__TSensor", "GetDJTAGAddress")

    Dim i As Long
    Dim i_tmp_str As String
    Dim i_ModCnt_lng As Long

    out_ModuleName_str = Split(in_ModuleName_str, ",")
    i_ModCnt_lng = UBound(out_ModuleName_str)
    ReDim out_DJTAGADDR_lng(i_ModCnt_lng)
    ReDim out_TestName_str(i_ModCnt_lng)
    ReDim out_EfuseItem_str(i_ModCnt_lng)

    For i = 0 To i_ModCnt_lng
        out_TestName_str(i) = "TS_" + out_ModuleName_str(i)
        Select Case out_ModuleName_str(i)
        Case "LCL0_FCM0"
            out_DJTAGADDR_lng(i) = &HFFF00080
            out_EfuseItem_str(i) = "TSR_FCM0"
        Case "LCL0_RMT_CPUB0"
            out_DJTAGADDR_lng(i) = &HFFF00080
            out_EfuseItem_str(i) = "TSR_RMT_CPUB0"
        Case "LCL0_RMT_CPUM0"
            out_DJTAGADDR_lng(i) = &HFFF00084
            out_EfuseItem_str(i) = "TSR_RMT_CPUM0"
        Case "LCL0_RMT_CPUM1"
            out_DJTAGADDR_lng(i) = &HFFF00084
            out_EfuseItem_str(i) = "TSR_RMT_CPUM1"
        Case "LCL0_RMT_CPUM2"
            out_DJTAGADDR_lng(i) = &HFFF00088
            out_EfuseItem_str(i) = "TSR_RMT_CPUM2"
        Case "LCL0_RMT_FCM0"
            out_DJTAGADDR_lng(i) = &HFFF00088
            out_EfuseItem_str(i) = "TSR_RMT_FCM0"
        Case "LCL0_RMT_FCM1"
            out_DJTAGADDR_lng(i) = &HFFF0008C
            out_EfuseItem_str(i) = "TSR_RMT_FCM1"
        Case "LCL0_RMT_DDRA"
            out_DJTAGADDR_lng(i) = &HFFF0008C
            out_EfuseItem_str(i) = "TSR_RMT_DDRA"
        Case "LCL1_MDM0"
            out_DJTAGADDR_lng(i) = &HFFF00180
            out_EfuseItem_str(i) = "TSR_MDM0"
        Case "LCL1_RMT_NCSI"
            out_DJTAGADDR_lng(i) = &HFFF00180
            out_EfuseItem_str(i) = "TSR_RMT_NCSI"
        Case "LCL1_RMT_NCHE"
            out_DJTAGADDR_lng(i) = &HFFF00184
            out_EfuseItem_str(i) = "TSR_RMT_NCHE"
        Case "LCL1_RMT_NPDT"
            out_DJTAGADDR_lng(i) = &HFFF00184
            out_EfuseItem_str(i) = "TSR_RMT_NPDT"
        Case "LCL1_RMT_5GTOP"
            out_DJTAGADDR_lng(i) = &HFFF00188
            out_EfuseItem_str(i) = "TSR_RMT_5GTOP"
        Case "LCL1_RMT_4GTOP"
            out_DJTAGADDR_lng(i) = &HFFF00188
            out_EfuseItem_str(i) = "TSR_RMT_4GTOP"
        Case "LCL2_G3D"
            out_DJTAGADDR_lng(i) = &HFFF00280
            out_EfuseItem_str(i) = "TSR_G3D"
        Case "LCL2_RMT_GPU0"
            out_DJTAGADDR_lng(i) = &HFFF00280
            out_EfuseItem_str(i) = "TSR_RMT_GPU0"
        Case "LCL2_RMT_GPU1"
            out_DJTAGADDR_lng(i) = &HFFF00284
            out_EfuseItem_str(i) = "TSR_RMT_GPU1"
        Case "LCL2_RMT_NPU0"
            out_DJTAGADDR_lng(i) = &HFFF00284
            out_EfuseItem_str(i) = "TSR_RMT_NPU0"
        Case "LCL2_RMT_ISP"
            out_DJTAGADDR_lng(i) = &HFFF00288
            out_EfuseItem_str(i) = "TSR_RMT_ISP"
        Case "LCL2_RMT_M1"
            out_DJTAGADDR_lng(i) = &HFFF00288
            out_EfuseItem_str(i) = "TSR_RMT_M1"
        Case "LCL2_RMT_M2"
            out_DJTAGADDR_lng(i) = &HFFF0028C
            out_EfuseItem_str(i) = "TSR_RMT_M2"
        Case "LCL2_RMT_DDRB"
            out_DJTAGADDR_lng(i) = &HFFF0028C
            out_EfuseItem_str(i) = "TSR_RMT_DDRB"

        Case Else
            MsgBox "TSENSOR_SOC_TEST ERROR: Unknown Module Name"
            Stop
        End Select
    Next i

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function TSensor_SOC_DSSC_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_CaptureSampleSize_lng As Long, _
       in_CaptureModuleName_str As String, _
       Optional FlagEfuseUpdate As Boolean = False, _
       Optional EfuseNameUpdate As String) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__TSensor", "TSensor_SOC_DSSC_vbt", TheExec.DataManager.InstanceName)

    '   Declarations
    Dim i As Long
    Dim io_ModCount As Long

    Dim Site As Variant

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_SignalName_str As String = "TSENSOR_SOC_ReadCode"

    Dim i_RunPrecondition_bool As Boolean

    Dim i_ModuleName_str() As String
    Dim i_TestName_str() As String
    Dim i_PinName_str() As String

    Dim i_Chuck_Temp_SDBL As New SiteDouble
    Dim i_Module_Count As New SiteLong

    Dim t_Temp_Ready_1_PLD As New PinListData
    Dim t_Temp_Mean_1_PLD As New PinListData
    Dim t_Temp_Delta_Mean_1_PLD As New PinListData
    Dim t_Temp_Delta_Min_1_PLD As New PinListData
    Dim t_Temp_Delta_Max_1_PLD As New PinListData

    '    Dim t_Temp_Ready_2_PLD As New PinListData
    '    Dim t_Temp_Mean_2_PLD As New PinListData
    '    Dim t_Temp_Delta_Mean_2_PLD As New PinListData
    '    Dim t_Temp_Delta_Min_2_PLD As New PinListData
    '    Dim t_Temp_Delta_Max_2_PLD As New PinListData


    '   I. ApplyLevelsTiming & prepare the Test Conditions
    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If

    Call GetCapturePins(in_CaptureModuleName_str, i_CapturePins_PL, i_ModuleName_str, i_TestName_str, i_PinName_str, io_ModCount)

    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites
            i_Chuck_Temp_SDBL = GetChuckTemp_dbl
        Next
    Else
        i_Chuck_Temp_SDBL = GlobalVariable_Chuck_Temp
    End If

    '   II. Setup DSSC
    '    TheHdw.Digital.Pins(i_CapturePins_PL).InitState = chInitoff    ' SUSPECT TO BE USELESS...
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng, _
         i_CapCodes_PLD)


    '   III. Begin the test
    If i_RunPrecondition_bool Then
        TheHdw.Patterns(in_ConfigPatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If

    If TheExec.TesterMode = testModeOnline Then
        TheHdw.Patterns(in_CapturePatName_PT).Start
        TheHdw.Digital.Patgen.HaltWait
    End If

    '   IV. Backgroud Calculation
    i_Module_Count = io_ModCount + 1

    rundsp.E_06_TSensor_SoC_Cal _
            i_CapCodes_PLD, _
            i_Chuck_Temp_SDBL, _
            i_Module_Count, _
            t_Temp_Ready_1_PLD, _
            t_Temp_Mean_1_PLD, _
            t_Temp_Delta_Mean_1_PLD, _
            t_Temp_Delta_Min_1_PLD, _
            t_Temp_Delta_Max_1_PLD
    '            t_Temp_Ready_2_PLD, _
                 '            t_Temp_Mean_2_PLD, _
                 '            t_Temp_Delta_Mean_2_PLD, _
                 '            t_Temp_Delta_Min_2_PLD, _
                 '            t_Temp_Delta_Max_2_PLD
    'every DSP has a potential risk of incomplete samples...

    '   V. Datalog

    ''    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Ready_1_PLD.pin(0), _
            PinName:=i_PinName_str(0), _
            forceResults:=tlForceFlow

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Mean_1_PLD.Pins(0), _
            PinName:=i_PinName_str(0), _
            forceResults:=tlForceFlow

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Delta_Mean_1_PLD.Pins(0), _
            PinName:=i_PinName_str(0), _
            forceResults:=tlForceFlow

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Delta_Min_1_PLD.Pins(0), _
            PinName:=i_PinName_str(0), _
            forceResults:=tlForceFlow

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Delta_Max_1_PLD.Pins(0), _
            PinName:=i_PinName_str(0), _
            forceResults:=tlForceFlow

    '    If i_Module_Count = 2 Then
    '        TheExec.Flow.TestLimit _
             '                ResultVal:=t_Temp_Ready_2_PLD.Pins(0), _
             '                TName:=i_TestName_str(1) + "_Ready", _
             '                PinName:=i_PinName_str(1), _
             '                ForceResults:=tlForceFlow
    '
    '        TheExec.Flow.TestLimit _
             '                ResultVal:=t_Temp_Mean_2_PLD.Pins(0), _
             '                TName:=i_TestName_str(1) + "_MEAN_TEMP", _
             '                PinName:=i_PinName_str(1), _
             '                ForceResults:=tlForceFlow
    '
    '        TheExec.Flow.TestLimit _
             '                ResultVal:=t_Temp_Delta_Mean_2_PLD.Pins(0), _
             '                TName:=i_TestName_str(1) + "_DELTA_MEAN", _
             '                PinName:=i_PinName_str(1), _
             '                ForceResults:=tlForceFlow
    '
    '        TheExec.Flow.TestLimit _
             '                ResultVal:=t_Temp_Delta_Min_2_PLD.Pins(0), _
             '                TName:=i_TestName_str(1) + "_DELTA_MIN", _
             '                PinName:=i_PinName_str(1), _
             '                ForceResults:=tlForceFlow
    '
    '        TheExec.Flow.TestLimit _
             '                ResultVal:=t_Temp_Delta_Max_2_PLD.Pins(0), _
             '                TName:=i_TestName_str(1) + "_DELTA_MAX", _
             '                PinName:=i_PinName_str(1), _
             '                ForceResults:=tlForceFlow
    '    End If

    ''    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019
    '   X.
'    If FlagEfuseUpdate = True Then
'        Call hiefuse.MULT_EFUSE_ITEM_UPDATE(EfuseNameUpdate, 0, Site, t_Temp_Mean_1_PLD.Pins(0), 0)
'    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Private Function GetCapturePins( _
        in_ModuleName_str As String, _
        out_CapturePins_PL As PinList, _
        ByRef io_ModuleName_str() As String, _
        ByRef io_TestName_str() As String, _
        ByRef io_PinName_str() As String, _
        ByRef i_ModCnt_lng As Long _
      ) As Long

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__TSensor", "GetCapturePins")

    '   Declarations
    Dim i As Long
    Dim i_CapturePins_str(2) As String
    Dim i_tmp_str As String

    io_ModuleName_str = Split(in_ModuleName_str, ",")
    i_ModCnt_lng = UBound(io_ModuleName_str)
    ReDim io_TestName_str(i_ModCnt_lng)
    ReDim io_PinName_str(i_ModCnt_lng)

    i_CapturePins_str(0) = "DigiCap_Tsensor_CPU_GPU"
    i_CapturePins_str(1) = "DigiCap_Tsensor_MODEM"
    i_CapturePins_str(2) = "DigiCap_Tsensor_CPU_GPU"

    i_tmp_str = ""
    For i = 0 To i_ModCnt_lng
        io_TestName_str(i) = "TS_" + io_ModuleName_str(i)
        Select Case io_ModuleName_str(i)
        Case "CPU"
            io_PinName_str(i) = i_CapturePins_str(0)
        Case "MDM"
            io_PinName_str(i) = i_CapturePins_str(1)
        Case "GPU"
            io_PinName_str(i) = i_CapturePins_str(2)
        Case Else
            MsgBox "TSENSOR_SOC_TEST ERROR: Unknown Module Name"
            Stop
        End Select
        If InStr(1, i_tmp_str, io_PinName_str(i), vbTextCompare) > 1 Then
            MsgBox "TSENSOR_SOC_TEST ERROR: Duplicate Module Name or Conflicted Pins"
            Stop
        Else
            i_tmp_str = i_tmp_str + "," + io_PinName_str(i)
        End If
    Next i

    If i_tmp_str = "" Then
        MsgBox "TSENSOR_SOC_TEST ERROR: Empty Module Name"
        Stop
    Else
        out_CapturePins_PL = MID$(i_tmp_str, 2)
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetChuckTemp_dbl() As Double
' this function need to be updated
    GetChuckTemp_dbl = 25#

End Function


Public Function TSENSOR_STLM75_Test(SlaveCount_lng As Long, _
                                    FlagEfuseUpdate As Boolean, EfuseNameUpdate As String, _
                                    Optional ConnectAllPins As Boolean = True, _
                                    Optional LoadLevels As Boolean = True, _
                                    Optional LoadTiming As Boolean = True, _
                                    Optional relayMode As tlRelayMode) As Long

    On Error GoTo errHandler

    TheHdw.Digital.Pins("TSENSOR_SCL,TSENSOR_SDA").Connect

    TheHdw.Digital.ApplyLevelsTiming ConnectAllPins, LoadLevels, LoadTiming, relayMode

    Dim i As Long
    Dim Temp_Mean As New SiteDouble
    Dim Temp_MaxDelta As New SiteDouble
    Dim Site As Variant
    Dim i_tmpReadData_SLNG As SiteLong
    Dim ReadData() As SiteLong
    ReDim ReadData(SlaveCount_lng - 1)

    Dim Temp_value() As New SiteDouble
    ReDim Temp_value(SlaveCount_lng - 1) As New SiteDouble


    For i = 0 To SlaveCount_lng - 1
        Call I2C_Read(i, i_tmpReadData_SLNG)
        Call I2C_Read(i, ReadData(i))
        For Each Site In TheExec.Sites
            If ReadData(i).ShiftRight(15) = 0 Then
                Temp_value(i) = ReadData(i).BitwiseAnd(&H7FFF).ShiftRight(7).Multiply(0.5)
            ElseIf ReadData(i).ShiftRight(15) = 1 Then
                Temp_value(i) = ReadData(i).BitwiseAnd(&H7FFF).ShiftRight(7).Multiply(-0.5)
            End If
        Next Site
        If TheExec.TesterMode = testModeOffline Then 'add for plus offline
            Temp_value(i) = -99
        End If
        TheExec.Flow.TestLimit ResultVal:=Temp_value(i), forceResults:=tlForceFlow
    Next i

    If SlaveCount_lng = 2 Then
        Temp_Mean = Temp_value(0).Add(Temp_value(1)).Divide(2)
        Temp_MaxDelta = Temp_value(0).Subtract(Temp_value(1)).Abs
    Else
        Temp_Mean = Temp_value(0)
        Temp_MaxDelta = Temp_value(0)
    End If

    TheExec.Flow.TestLimit ResultVal:=Temp_Mean, forceResults:=tlForceFlow
    TheExec.Flow.TestLimit ResultVal:=Temp_MaxDelta, forceResults:=tlForceFlow

    '    TheHdw.Protocol.Ports("TS_I2C_Pins").Enabled = False

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function I2C_Read(ByVal SlaveAdr As Long, ByRef mdata As SiteLong) As Long
    On Error GoTo errHandler

    Dim vsite As Variant
    Dim ExicutionType As tlNWireExecutionType
    Dim myPLD As New PinListData
    Dim myDSPwave As New DSPWave

    Set mdata = New SiteLong  ''?

    TheHdw.Protocol.Ports("TS_I2C_Pins").Enabled = True

    If TheHdw.Protocol.Ports("TS_I2C_Pins").ModuleFiles.IsLoading Then
        ExicutionType = tlNWireExecutionType_PushToStack
    Else
        With TheHdw.Protocol.Ports("TS_I2C_Pins").NWire.CMEM
            .MoveMode = tlNWireCMEMMoveMode_Databus
            myPLD = .DSPWave
            myDSPwave = myPLD.Pins("TS_I2C_Pins")
        End With
        ''        ExicutionType = tlNWireExecutionType_CaptureInCMEM
    End If

    With TheHdw.Protocol.Ports("TS_I2C_Pins")
        With .NWire.Frames("I2CRead")
            .Fields("A1A0").Value = SlaveAdr
            .Fields("AddrByte").Value = 0
            .Fields("A1A0_Read").Value = SlaveAdr
            .Execute tlNWireExecutionType_CaptureInCMEM    '     ExicutionType
        End With
        Call .IdleWait
    End With

    If TheHdw.Protocol.Ports("TS_I2C_Pins").ModuleFiles.IsLoading Then    ' check pass/fail if not recording
        mdata = 0
        Exit Function
    End If

    For Each vsite In TheExec.Sites
        If TheExec.TesterMode = testModeOnline Then
            If TheHdw.Protocol.Ports("TS_I2C_Pins").NWire.CMEM.Transactions.count = 0 Then    ' no reading, sth wrong??
                mdata = -1
                TheExec.AddOutput "Error! Instance <" + TheExec.DataManager.InstanceName + "> has abnormal PA reading, check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
                '            TheExec.AddOutput "Address: &H" + Hex(Addr_8_bits), vbRed
            Else
                mdata = myDSPwave.Element(0)
            End If
        End If
    Next vsite

    '''    I2C_Read = mData

    TheHdw.Protocol.Ports("TS_I2C_Pins").Enabled = False

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetGetHandlerTemperature(Optional FlagEfuseUpdate As Boolean = False, _
                                         Optional EfuseNameUpdate As String = "TSR_REF") As Long
    On Error GoTo errHandler

    Dim Overlap As Boolean
    Dim Num_Read As Integer
    Dim Idx As Integer
    Dim i As Integer
    Dim Realtime_temp As New SiteDouble
    Dim Site As Variant

    Dim ArmSettingTemp As Double
    If TheExec.RunMode = runModeProduction Then
        ArmSettingTemp = setTemp
    Else
        ArmSettingTemp = JobName2Temp    'GetproberTemp.ChuckSettingTemp  '.ArmSettingTemp
    End If

    Dim ArmCurrentTemp As New SiteDouble

    For Each Site In TheExec.Sites
        If TheExec.RunMode = runModeProduction Then
            ArmCurrentTemp(Site) = currTemp
        Else
            ArmCurrentTemp(Site) = JobName2Temp    'GetproberTemp.ChuckCurrentTemp(site)  '.ArmCurrentTemp(site)
        End If

        GlobalVariable_Chuck_Temp = ArmCurrentTemp(Site)
    Next Site

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    TheExec.Flow.TestLimit ResultVal:=ArmSettingTemp, unit:=unitCustom, customUnit:="C", forceResults:=tlForceFlow, TName:="SET_TEMPERATURE"

    TheExec.Flow.TestLimit ResultVal:=ArmCurrentTemp, unit:=unitCustom, customUnit:="C", forceResults:=tlForceFlow, TName:="READ_TEMPERATURE"

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    For Each Site In TheExec.Sites
        If FlagEfuseUpdate = True Then
            Call hiefuse.MULT_EFUSE_ITEM_UPDATE(EfuseNameUpdate, 0, Site, GlobalVariable_Chuck_Temp, 0)
            'For DataMonitor
            gTsrREFValue = GlobalVariable_Chuck_Temp
        End If
    Next Site

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function JobName2Temp() As Double

    Dim JobName As String
    JobName = TheExec.CurrentJob

    Dim JobList() As String
    JobList = Split(JobName, "_")

    Dim i As Integer
    Dim j As Integer
    Dim ascii As Integer
    Dim TempIndex As Integer
    Dim Sign As Integer
    TempIndex = -1

    For i = UBound(JobList) To 0 Step -1
        ascii = Asc(Right(JobList(i), 1))   'judge if last character is "c" or "C"
        If ascii = 67 Or ascii = 99 Then
            ascii = Asc(Left(JobList(i), 1))    'judge if first character is "C" or figure
            If ascii >= 48 And ascii <= 57 Then     'judge if middle character is figure
                Sign = 1
                For j = 1 To Len(JobList(i)) - 1 Step 1
                    ascii = Asc(MID(JobList(i), j, 1))
                    If ascii >= 48 And ascii <= 57 Then
                        TempIndex = i
                    Else
                        TempIndex = -1
                        GoTo Loop_i
                    End If
                Next j
            ElseIf ascii = 76 Or ascii = 108 Then
                Sign = -1
                Dim Length_Temp As Long
                Length_Temp = Len(JobList(i)) - 1
                JobList(i) = MID(JobList(i), 2, Length_Temp)
                For j = 1 To Len(JobList(i)) - 1 Step 1
                    ascii = Asc(MID(JobList(i), j, 1))
                    If ascii >= 48 And ascii <= 57 Then
                        TempIndex = i
                    Else
                        TempIndex = -1
                        GoTo Loop_i
                    End If
                Next j

            Else
                TempIndex = -1
                GoTo Loop_i
            End If
        Else
            TempIndex = -1
            GoTo Loop_i
        End If
        If TempIndex <> -1 Then Exit For
Loop_i:
    Next i

    If TempIndex <> -1 Then
        Dim Length As Long
        Length = Len(JobList(TempIndex)) - 1

        Dim tempstr As String
        tempstr = Left(JobList(TempIndex), Length)

        JobName2Temp = CDbl(tempstr) * Sign
    Else
        JobName2Temp = 27
    End If

End Function
'
'Public Function TSensor_DJTAG_PA_vbt( _
'       in_RelayMode_tl As tlRelayMode, _
'       in_CaptureSampleSize_lng As Long, _
'       in_CaptureModuleName_str As String, _
'       Optional FlagEfuseUpdate As Boolean = False) As Long
'
'    On Error GoTo errHandler
'    Call LogCalledFunctions("VBT__TSensor", "TSensor_DJTAG_PA_vbt", TheExec.DataManager.InstanceName)
'
'    '   Declarations
'    Dim i As Long
'
'    Dim Site As Variant
'
'    '    Dim i_CapCodes_PLD As New PinListData
'    Dim i_CapCodes_DSP As New DSPWave
'    Dim i_CapturePins_PL As New PinList
'    Const i_SignalName_str As String = "TSENSOR_DJTAG_ReadCode_PA"
'
'    Dim i_ModuleName_str() As String
'    Dim i_TestName_str() As String
'    Dim i_TestCount_lng As Long
'
'    Dim i_Chuck_Temp_SDBL As New SiteDouble
'
'    Dim i_Temp_Mean_PLD As New PinListData
'    Dim t_Temp_Mean_DSP As New DSPWave
'    Dim i_Temp_Delta_Mean_PLD As New PinListData
'    Dim t_Temp_Delta_Mean_DSP As New DSPWave
'
'    Dim i_EfuseItem_str() As String
'    Dim i_EfuseAddr_lng() As Long
'
'
'    '   I. ApplyLevelsTiming & prepare the Test Conditions
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, in_RelayMode_tl
'
'    'add PA clock start on 20191125 in SJ
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'    TheHdw.Protocol.Ports("CLK_38M4").Enabled = True
'    TheHdw.Protocol.Ports("CLK_38M4").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_38M4").IdleWait
'    TheHdw.Wait 0.001 '* 10
'
'
'    Call GetDJTAGAddress(in_CaptureModuleName_str, i_ModuleName_str, i_TestName_str, i_EfuseItem_str, i_EfuseAddr_lng)
'
'    i_TestCount_lng = UBound(i_ModuleName_str)
'    If i_TestCount_lng + 1 <> in_CaptureSampleSize_lng Then
'        MsgBox "CaptureSampleSize should equal to number of Modules, check Instance!"
'    End If
'
'    i_CapCodes_DSP.CreateConstant &H19F, i_TestCount_lng + 1, DspLong
'    i_Temp_Mean_PLD.AddPin ("JTAG_TDO")
'    i_Temp_Delta_Mean_PLD.AddPin ("JTAG_TDO")
'    i_CapturePins_PL = "JTAG_TDO"
'
'    If TheExec.TesterMode = testModeOffline Then
'        For Each Site In TheExec.Sites
'            i_Chuck_Temp_SDBL = GetChuckTemp_dbl
'        Next
'    Else
'        i_Chuck_Temp_SDBL = GlobalVariable_Chuck_Temp
'    End If
'
'
'    '   II. Setup vector using initial state
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''SETUP''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    TheHdw.Digital.Pins("DFT_EN,JTAG_MODE,JTAG_SEL0,JTAG_TDI,JTAG_TMS_SWDIO,GPIO_045,GPIO_046,GPIO_047,JTAG_TCK_SWCLK").InitState = chInitLo
'    TheHdw.Digital.Pins("JTAG_SEL1,JTAG_TRST_N,PMU_RST_SOC_N,TEST_MODE,GPIO_043,GPIO_044").InitState = chInitHi
'    TheHdw.Digital.Pins("BOOT_MODE,JTAG_TDO").InitState = chInitoff
'    TheHdw.Wait 62.5 * US   ' repeat 600 *104.167ns
'    TheHdw.Digital.Pins("JTAG_TRST_N,PMU_RST_SOC_N").InitState = chInitLo
'    TheHdw.Wait 900 * US    'repeat 144*2 * loop 30 * 104.167ns
'    TheHdw.Digital.Pins("JTAG_TRST_N,PMU_RST_SOC_N").InitState = chInitHi
'    TheHdw.Wait 6 * mS    'repeat 144*2 * loop 200 * 104.167ns
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''SETUP''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'    '   III. Begin the test and get captured data using PA
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''PA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '   III.a Initial
'    TheHdw.Wait 10 * mS    'register file request
'
'    '    TheHdw.Digital.Pins("").InitState = chInitLo
'    '    TheHdw.Digital.Pins("").InitState = chInitHi
'
'    '   III.b PA port enable & capture setup
'    TheHdw.Protocol.Ports("JTAG_PINS").Enabled = True
'    TheHdw.Protocol.ModuleRecordingEnabled = True
'
'    Dim myPLD As New PinListData
'    Dim myDSPwave As New DSPWave
'    myDSPwave.CreateConstant &H0, 0, DspLong
'
'    With TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM
'        .MoveMode = tlNWireCMEMMoveMode_Databus
'        myPLD = .DSPWave
'        myDSPwave = myPLD.Pins("JTAG_Pins")
'    End With
'
'    '   III.c excute PA
'    If TheHdw.Protocol.Ports("JTAG_PINS").Modules.IsRecorded("DJTAG_Module", True, True) = False Then
'
'        Call DJTAG_WRITE("FFF00010", "00004071")
'        Call DJTAG_WRITE("FFF00110", "00004051")
'        Call DJTAG_WRITE("FFF00210", "00004071")
'
'        TheHdw.Wait 7 * mS    'register file request
'
'        Call DJTAG_READ("FFF00080")
'        Call DJTAG_READ("FFF00084")
'        Call DJTAG_READ("FFF00088")
'        Call DJTAG_READ("FFF0008C")
'        Call DJTAG_READ("FFF00180")
'        Call DJTAG_READ("FFF00184")
'        Call DJTAG_READ("FFF00188")
'        Call DJTAG_READ("FFF00280")
'        Call DJTAG_READ("FFF00284")
'        Call DJTAG_READ("FFF00288")
'        Call DJTAG_READ("FFF0028C")
'
'        TheHdw.Protocol.Ports("JTAG_PINS").Modules.StopRecording
'    End If
'
'    With TheHdw.Protocol.Ports("JTAG_Pins")
'        .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
'        myPLD = .NWire.CMEM.DSPWave
'        myDSPwave = myPLD.Pins("JTAG_Pins")
'        .Modules("DJTAG_Module").Start
'        .IdleWait
'    End With
'
'    '   III.d get the captured data
'    For Each Site In TheExec.Sites
'        If TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM.Transactions.count = 0 Then    ' no reading, sth wrong??
'            TheExec.AddOutput "Error! Instance <" + TheExec.DataManager.InstanceName + "> has abnormal PA reading, check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
'        Else
'            For i = 0 To i_TestCount_lng
'                '''                    i_CapCodes_DSP(Site).Element(i) = myDSPwave(Site).Element(i * 2)    '''Each PA read captured 34 bits(2 * 32-bit element )
'                If i Mod 2 = 0 Then
'                    i_CapCodes_DSP(Site).Element(i) = myDSPwave(Site).Element(i) And &HFFFF&    '''Each PA read captured 34 bits(2 * 32-bit element )
'                Else
'                    i_CapCodes_DSP(Site).Element(i) = myDSPwave(Site).Element(i - 1) / 2 ^ 16    '''Each PA read captured 34 bits(2 * 32-bit element )
'                End If
'            Next i
'        End If
'    Next Site
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''PA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '''    '   IV. Calculation
'    '''    t_Temp_Mean_DSP = i_CapCodes_DSP.BitwiseExtract(16, 0).Subtract(414).ConvertDataTypeTo(DspDouble)
'    '''    t_Temp_Mean_DSP = t_Temp_Mean_DSP.Multiply(165 / (700 - 414)).Subtract(40#)
'    '''
'    '''    t_Temp_Delta_Mean_DSP = t_Temp_Mean_DSP.Subtract(currTemp)
'
'    '   IV. Backgroud Calculation
'    rundsp.E_06_TSensor_DJTAG_Cal _
'            i_CapCodes_DSP, _
'            i_Chuck_Temp_SDBL, _
'            i_Temp_Mean_PLD, _
'            i_Temp_Delta_Mean_PLD
'    'every DSP has a potential risk of incomplete samples...
'    For Each Site In TheExec.Sites
'        t_Temp_Mean_DSP = i_Temp_Mean_PLD.Pins(i_CapturePins_PL).Value
'        t_Temp_Delta_Mean_DSP = i_Temp_Delta_Mean_PLD.Pins(i_CapturePins_PL).Value
'    Next Site
'
'    '   V. Datalog
'
'    TheExec.Flow.Limits.key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019
'
'    For i = 0 To i_TestCount_lng
'
'        TheExec.Flow.TestLimit _
'                resultVal:=t_Temp_Mean_DSP.Element(i), _
'                TName:="TS_MEAN_" + i_ModuleName_str(i), _
'                PinName:=i_CapturePins_PL, _
'                ForceResults:=tlForceFlow
'
'        TheExec.Flow.TestLimit _
'                resultVal:=t_Temp_Delta_Mean_DSP.Element(i), _
'                TName:="TS_DELTA_" + i_ModuleName_str(i), _
'                PinName:=i_CapturePins_PL, _
'                ForceResults:=tlForceFlow
'
'    Next i
'
'    TheExec.Flow.Limits.key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019
'    '   X.
'
'    'Update EFUSE flag
'    For Each Site In TheExec.Sites
'        For i = 0 To i_TestCount_lng
'            If FlagEfuseUpdate = True Then
'                Call HiEFUSE.MULT_EFUSE_ITEM_UPDATE(i_EfuseItem_str(i), 0, Site, t_Temp_Mean_DSP.Element(i), 0)
'            End If
'        Next i
'    Next Site
'
'    'add PA clock stop on 20191125 in SJ
'    TheHdw.Protocol.Ports("CLK_32K").Enabled = True
'    TheHdw.Protocol.Ports("CLK_32K").NWire.Frames("Start_CLOCK").Execute
'    TheHdw.Protocol.Ports("CLK_32K").IdleWait
'    TheHdw.Wait 0.001
'
'    Exit Function
'errHandler:
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'Public Function DJTAG_READ(addr As String) As Long
''''''1 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''2 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "00")  -- r0xFFF3000C
'    Call JTAG_SDR_LONG_DJTAG("0" + addr, "", 34)
'
'    '''''3 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''4 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "01")  -- r0xFFF3000C
'    Call JTAG_SDR_LONG_DJTAG("2" + addr, "", 34)
'
'    '''''5 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''6 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "xx")  -- writePin(pin_tdo,"1111111111111111111111111111111100");
'    Call JTAG_SDR_LONG_DJTAG("000000000", "", 34)
'
'    '''''7 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''8 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "00")  -- r0xFFF3000C
'    Call JTAG_SDR_LONG_DJTAG("000000000", "000000000", 34)
'
'End Function
'
'Public Function DJTAG_WRITE(addr As String, Data As String) As Long
''''''1 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''2 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "10")  -- r0xFFF3000C
'    Call JTAG_SDR_LONG_DJTAG("1" + addr, "", 34)
'
'    '''''3 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''4 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "11")  -- r0xFFF3000C
'    Call JTAG_SDR_LONG_DJTAG("3" + addr, "", 34)
'
'    '''''5 CALL idleToShiftIR and CALL writeTDIandGoIdle(tdi_ir)  -- 10100111
'    Call JTAG_SIR(Bin2Hex("10100111"))
'    '''''6 CALL idleToShiftDR and CALL writeTDIandGoIdle(tdi_addr + "10")  -- writePin(pin_tdo,"1111111111111111111111111111111100");
'    Call JTAG_SDR_LONG_DJTAG("1" + Data, "", 34)
'
'End Function

Public Function Tsensor_GLITCH_vbt() As Long

    Dim t_TempMax_sdb As New SiteDouble
    Dim t_TempMin_sdb As New SiteDouble
    Dim t_TempMean_sdb As New SiteDouble
    Dim t_Temp_Gap_sdb As New SiteDouble

    On Error GoTo errHandler
    
    rundsp.E_06_Tsensor_Glitch Global_Djtag_Mean_DSP, _
                               t_TempMean_sdb, _
                               t_TempMin_sdb, _
                               t_TempMax_sdb, _
                               t_Temp_Gap_sdb
    TheExec.Flow.TestLimit _
            ResultVal:=t_TempMax_sdb, _
            TName:="TS_MODULE_MAX"

    TheExec.Flow.TestLimit _
            ResultVal:=t_TempMin_sdb, _
            TName:="TS_MODULE_MIN"

    TheExec.Flow.TestLimit _
            ResultVal:=t_TempMean_sdb, _
            TName:="TS_MODULE_MEAN"

    TheExec.Flow.TestLimit _
            ResultVal:=t_Temp_Gap_sdb, _
            TName:="TS_MAX_MEAN_GAP", _
            forceResults:=tlForceFlow
            
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Tsensor_Trim_Check_vbt() As Long

    Dim Site As Variant
    Dim EfuseNameRead_Temp_db As Double
    Dim Temp_Trimed_db As Double
    Dim i As Long
    Dim DJTAG_MODULE_str() As String

    ReDim DJTAG_MODULE_str(21)
    
    On Error GoTo errHandler

    'DJTAG MODULE TABLE

    DJTAG_MODULE_str(0) = "TSR_FCM0"
    DJTAG_MODULE_str(1) = "TSR_RMT_CPUB0"
    DJTAG_MODULE_str(2) = "TSR_RMT_CPUM0"
    DJTAG_MODULE_str(3) = "TSR_RMT_CPUM1"
    DJTAG_MODULE_str(4) = "TSR_RMT_CPUM2"
    DJTAG_MODULE_str(5) = "TSR_RMT_FCM0"
    DJTAG_MODULE_str(6) = "TSR_RMT_FCM1"
    DJTAG_MODULE_str(7) = "TSR_RMT_DDRa"
    DJTAG_MODULE_str(8) = "TSR_MDM0"
    DJTAG_MODULE_str(9) = "TSR_RMT_ncsi"
    DJTAG_MODULE_str(10) = "TSR_RMT_nche"
    DJTAG_MODULE_str(11) = "TSR_RMT_npdt"
    DJTAG_MODULE_str(12) = "TSR_RMT_5GTOP"
    DJTAG_MODULE_str(13) = "TSR_RMT_4GTOP"
    DJTAG_MODULE_str(14) = "TSR_G3D"
    DJTAG_MODULE_str(15) = "TSR_RMT_GPU0"
    DJTAG_MODULE_str(16) = "TSR_RMT_GPU1"
    DJTAG_MODULE_str(17) = "TSR_RMT_NPU0"
    DJTAG_MODULE_str(18) = "TSR_RMT_ISP"
    DJTAG_MODULE_str(19) = "TSR_RMT_M1"
    DJTAG_MODULE_str(20) = "TSR_RMT_M2"
    DJTAG_MODULE_str(21) = "TSR_RMT_DDRb"


    For i = 0 To 21
        For Each Site In TheExec.Sites
            If TheExec.TesterMode = testModeOnline Then
                EfuseNameRead_Temp_db = hiefuse.GET_MULT_EFUSE_ITEM_CODE_RD(DJTAG_MODULE_str(i), 0, CLng(Site))
            Else
                EfuseNameRead_Temp_db = 25
            End If
            'initialize
            Temp_Trimed_db = 0#
            If (EfuseNameRead_Temp_db > 128) Then
                Temp_Trimed_db = (128 - EfuseNameRead_Temp_db) / 4
            Else
                Temp_Trimed_db = EfuseNameRead_Temp_db / 4
            End If
            
            TheExec.Flow.TestLimitIndex = 2 * i
            
            TheExec.Flow.TestLimit _
                    ResultVal:=Global_Djtag_Mean_DSP(Site).Subtract(Temp_Trimed_db).Element(i), _
                    forceResults:=tlForceFlow
                    
            TheExec.Flow.TestLimitIndex = 2 * i + 1
            TheExec.Flow.TestLimit _
                    ResultVal:=Global_Djtag_Mean_DSP.Subtract(Temp_Trimed_db).Element(i) - GlobalVariable_Chuck_Temp, _
                    forceResults:=tlForceFlow
        Next Site
    Next i
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetChuckTemp(ct As Double, st As Double) As Long
    On Error GoTo errHandler
    If TheExec.RunMode = runModeDebug Then
        currTemp = 999
        setTemp = 999
    Else
        currTemp = ct
        setTemp = st
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


