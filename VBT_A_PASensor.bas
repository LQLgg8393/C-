Attribute VB_Name = "VBT_A_PASensor"
Option Explicit
Global g_PASensor_RegName_str() As String
Global g_PASensor_ValidRegName_str() As String


Public Type Info
    MinuendRegisterData As New SiteLong
    MinuendRegisterName As String
    SubtrahendRegisterData As New SiteLong
    SubtrahendRegisterName As String
    EfuseData As New SiteLong
End Type

Public Type AREA
    HV() As Info
    MV() As Info
    LV() As Info
End Type
Public PASensor_AREA As AREA
Private Const m_SETTLEWAIT_dbl As Double = 0.01   ' new spec settling time in sec,


Public Function PASensor_ReadCode_vbt( _
       in_RelayMode_tl As tlRelayMode, _
       in_ConfigPatName_PT As Pattern, _
       in_CapturePatName_PT As Pattern, _
       in_CaptureSampleSize_lng As Long, _
       in_SpecName_str As String, _
       in_TestConditionList_str As String, _
       Optional in_EfuseEnFlag As Boolean, _
       Optional in_EFUSEMinuend_str As String, _
       Optional in_EFUSESubtrahend_str As String) As Long
' EDITFORMAT1 1,RelayMode,tlRelayMode,,UnPowered will reset the DUT,in_RelayMode_tl|2,ConfigPatName_PT,Pattern,,,in_ConfigPatName_PT|3,CapturePatName_PT,Pattern,,,in_CapturePatName_PT|4,CaptureSampleSize_lng,Long,,,in_CaptureSampleSize_lng|5,RegisterNames_str,String,,comma separated list,in_RegisterNames_str|6,ValidRegisterNames_str,String,,comma separated list,in_ValidRegisterNames_str|7,SpecName_str,String,,,in_SpecName_str|8,TestConditionList_str,String,,comma separated list,in_TestConditionList_str

    On Error GoTo errHandler
    Call LogCalledFunctions("VBT__PASensor", "PASensor_ReadCode_vbt", TheExec.DataManager.InstanceName)

    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim Site As Variant

    Dim i_CapCodes_PLD As New PinListData
    Dim i_CapturePins_PL As New PinList
    Const i_SignalName_str As String = "PASensor_ReadCode"

    Dim i_RunPrecondition_bool As Boolean
    Dim i_CaptureValid_bool As Boolean

    Dim i_TestCount_lng                     As Long
    Dim i_TestCondition_str()               As String
    Dim i_SpecNameCount_lng                 As Long
    Dim i_SpecNameList_str()                As String
    Dim i_EFUSEMinuend_str()                As String
    Dim i_EFUSESubtrahend_str()             As String
    Dim i_EFUSEMinuend_lng                  As Long
    Dim i_EFUSESubtrahend_lng               As Long
    
    Dim i_BaseSpec_SDBL                     As New SiteDouble
    Dim i_InitialSpecVal()                  As New SiteDouble
    Dim i_TestSpec_dbl()                    As Double
    
    Dim i_EFUSEMinuend_dbl()                As SiteDouble
    Dim i_EFUSESubtrahend_dbl()             As SiteDouble
    Dim i_RegData_PLD                       As New PinListData
    Dim i_ValidRegData_PLD                  As New PinListData
    Dim t_RegData_DSP                       As New DSPWave
    Dim t_ValidRegData_DSP                  As New DSPWave


    '   I. ApplyLevelsTiming and setup DSSC
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    If nonblank(in_ConfigPatName_PT) Then
        i_RunPrecondition_bool = True
    End If
 
    i_SpecNameList_str = Split(in_SpecName_str, ",")
    i_SpecNameCount_lng = UBound(i_SpecNameList_str)
    
    i_TestCondition_str = Split(in_TestConditionList_str, ",")
    i_TestCount_lng = UBound(i_TestCondition_str)
    ReDim i_TestSpec_dbl(i_TestCount_lng)

    ReDim i_InitialSpecVal(i_SpecNameCount_lng)
    For i = 0 To i_SpecNameCount_lng
        i_InitialSpecVal(i) = TheExec.Specs.Find(i_SpecNameList_str(i)).CurrentValue
    Next i
    
    For i = 0 To i_TestCount_lng
        i_TestSpec_dbl(i) = LookupSpecTable("PA_" + i_TestCondition_str(i))
    Next i
    
    If in_EFUSEMinuend_str <> "" And in_EFUSESubtrahend_str <> "" Then
        i_EFUSEMinuend_str = Split(Trim(in_EFUSEMinuend_str), ",")
        i_EFUSESubtrahend_str = Split(Trim(in_EFUSESubtrahend_str), ",")
        i_EFUSEMinuend_lng = UBound(i_EFUSEMinuend_str)
        i_EFUSESubtrahend_lng = UBound(i_EFUSESubtrahend_str)
        If i_EFUSEMinuend_lng <> i_EFUSESubtrahend_lng Then
            TheExec.Datalog.WriteComment " Please check in_EFUSEMinuend_str and in_EFUSESubtrahend_str"
            Exit Function
        End If
        For i = 0 To i_TestCount_lng
            If i_TestCondition_str(i) = "0P8" Then
                ReDim PASensor_AREA.HV(i_EFUSEMinuend_lng)
                For j = 0 To i_EFUSEMinuend_lng
                    PASensor_AREA.HV(j).MinuendRegisterName = i_EFUSEMinuend_str(j)
                    PASensor_AREA.HV(j).SubtrahendRegisterName = i_EFUSESubtrahend_str(j)
                Next j
            ElseIf i_TestCondition_str(i) = "0P7" Then
                ReDim PASensor_AREA.MV(i_EFUSEMinuend_lng)
                For j = 0 To i_EFUSEMinuend_lng
                    PASensor_AREA.MV(j).MinuendRegisterName = i_EFUSEMinuend_str(j)
                    PASensor_AREA.MV(j).SubtrahendRegisterName = i_EFUSESubtrahend_str(j)
                Next j
            ElseIf i_TestCondition_str(i) = "0P6" Then
                ReDim PASensor_AREA.LV(i_EFUSEMinuend_lng)
                For j = 0 To i_EFUSEMinuend_lng
                    PASensor_AREA.LV(j).MinuendRegisterName = i_EFUSEMinuend_str(j)
                    PASensor_AREA.LV(j).SubtrahendRegisterName = i_EFUSESubtrahend_str(j)
                Next j
            Else
                TheExec.Datalog.WriteComment " Please check Test Condition"
                Exit Function
            End If
        Next i
    End If
    
    '   II. prepare the Test Conditions
    i_CapturePins_PL = "JTAG_TDO"
    Call htl_SetupDSSC( _
         in_CapturePatName_PT, _
         i_CapturePins_PL, _
         i_SignalName_str, _
         in_CaptureSampleSize_lng * 37, _
         i_CapCodes_PLD)


    '   III. Looping the test
    For i = 0 To i_TestCount_lng
        'ApplyUniformSpecToHW
        For j = 0 To i_SpecNameCount_lng
            TheExec.Overlays.ApplyUniformSpecToHW i_SpecNameList_str(j), i_TestSpec_dbl(i), True, True
        Next j
        
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
    For Each Site In TheExec.Sites
        For i = 0 To i_SpecNameCount_lng
            TheExec.Overlays.ApplyUniformSpecToHW i_SpecNameList_str(i), i_InitialSpecVal(i), True, True
        Next i
    Next Site
        
    '   Check Init_PASensor_Array
    Call Init_PASensor_Array

    '   IV. Backgroud Calculation
    rundsp.E_02_PASENSOR_Cal i_CapCodes_PLD, i_TestCount_lng, i_RegData_PLD, i_ValidRegData_PLD
 
    t_RegData_DSP = i_RegData_PLD.Pins(i_CapturePins_PL)
    t_ValidRegData_DSP = i_ValidRegData_PLD.Pins(i_CapturePins_PL)
    
    '   V. Datalog

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName    ' Added for Opening Nonsequential Execution, William LIAO, 10/10/2019

    For i = 0 To i_TestCount_lng
        For j = 0 To in_CaptureSampleSize_lng * 4 - 1
            TheExec.Flow.TestLimit _
                    ResultVal:=t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j), _
                    TName:=g_PASensor_RegName_str(j) + "_" + i_TestCondition_str(i), _
                    PinName:=i_CapturePins_PL, _
                    forceResults:=tlForceFlow
             
            ' ================================================================================
            '                Get Info from PASensor DSPWAVE
            ' ================================================================================
            If in_EFUSEMinuend_str <> "" And in_EFUSESubtrahend_str <> "" Then
                If i_TestCondition_str(i) = "0P6" Then
                    For m = 0 To UBound(PASensor_AREA.LV)
                        If UCase(PASensor_AREA.LV(m).MinuendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.LV(m).MinuendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        ElseIf UCase(PASensor_AREA.HV(m).SubtrahendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.LV(m).SubtrahendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        End If
                    Next m
                ElseIf i_TestCondition_str(i) = "0P7" Then
                    For m = 0 To UBound(PASensor_AREA.MV)
                        If UCase(PASensor_AREA.MV(m).MinuendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.MV(m).MinuendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        ElseIf UCase(PASensor_AREA.HV(m).SubtrahendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.MV(m).SubtrahendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        End If
                    Next m
                ElseIf i_TestCondition_str(i) = "0P8" Then
                    For m = 0 To UBound(PASensor_AREA.HV)
                        If UCase(PASensor_AREA.HV(m).MinuendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.HV(m).MinuendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        ElseIf UCase(PASensor_AREA.HV(m).SubtrahendRegisterName) = UCase(g_PASensor_RegName_str(j)) Then
                            PASensor_AREA.HV(m).SubtrahendRegisterData = t_RegData_DSP.Element(i * in_CaptureSampleSize_lng * 4 + j)
                            Exit For
                        End If
                    Next m
                Else
                    TheExec.Datalog.WriteComment " Please check Test Condition"
                    Exit Function
                End If
             End If
        Next j
        
        For j = 0 To in_CaptureSampleSize_lng - 1
            TheExec.Flow.TestLimit _
                    ResultVal:=t_ValidRegData_DSP.Element(i * in_CaptureSampleSize_lng + j), _
                    TName:=g_PASensor_ValidRegName_str(j) + "_" + i_TestCondition_str(i), _
                    PinName:=i_CapturePins_PL, _
                    forceResults:=tlForceFlow
        Next j
    Next i

    TheExec.Flow.Limits.Key = tlFlowLimitsKeyNone    ' Added for Closing Nonsequential Execution, William LIAO, 10/10/2019

    ' ================================================================================
    '                Update Efuse Table and Trim Data Init
    ' ================================================================================

    If in_EfuseEnFlag = True And in_EFUSEMinuend_str <> "" And in_EFUSESubtrahend_str <> "" Then
        For i = 0 To i_TestCount_lng
            If i_TestCondition_str(i) = "0P6" Then
                If i_EFUSEMinuend_lng + 1 = 9 And i_EFUSESubtrahend_lng + 1 = 9 Then
                    For Each Site In TheExec.Sites
                        PASensor_AREA.LV(2).EfuseData(Site) = PASensor_AREA.LV(2).MinuendRegisterData(Site) - PASensor_AREA.LV(2).SubtrahendRegisterData(Site)
                        PASensor_AREA.LV(3).EfuseData(Site) = PASensor_AREA.LV(3).MinuendRegisterData(Site) - PASensor_AREA.LV(3).SubtrahendRegisterData(Site)
                        If PASensor_AREA.LV(2).EfuseData(Site) < -15 Or PASensor_AREA.LV(2).EfuseData(Site) > 15 Or PASensor_AREA.LV(3).EfuseData(Site) < -15 Or PASensor_AREA.LV(3).EfuseData(Site) > 15 Then
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.LV(2).EfuseData, TName:="Delta_G3D_ULVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.LV(3).EfuseData, TName:="Delta_G3D_LVT"
                        Else
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("G3D_ULVT", 0, Site, PASensor_AREA.LV(2).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("G3D_LVT", 0, Site, PASensor_AREA.LV(3).EfuseData(Site), 0, 0)
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("G3D_ULVT")).gDataValue(Site) = CDbl(PASensor_AREA.LV(2).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("G3D_LVT")).gDataValue(Site) = CDbl(PASensor_AREA.LV(3).EfuseData(Site))
                        End If
                    Next Site
                Else
                    TheExec.Datalog.WriteComment "ERROR, don't support the size of Minuend or Subtrahend "
                    Exit Function
                End If
             ElseIf i_TestCondition_str(i) = "0P7" Then
                If i_EFUSEMinuend_lng + 1 = 9 And i_EFUSESubtrahend_lng + 1 = 9 Then
                    For Each Site In TheExec.Sites
                        PASensor_AREA.MV(0).EfuseData(Site) = PASensor_AREA.MV(0).MinuendRegisterData(Site) - PASensor_AREA.MV(0).SubtrahendRegisterData(Site)
                        PASensor_AREA.MV(1).EfuseData(Site) = PASensor_AREA.MV(1).MinuendRegisterData(Site) - PASensor_AREA.MV(1).SubtrahendRegisterData(Site)
                        PASensor_AREA.MV(4).EfuseData(Site) = PASensor_AREA.MV(4).MinuendRegisterData(Site) - PASensor_AREA.MV(4).SubtrahendRegisterData(Site)
                        PASensor_AREA.MV(5).EfuseData(Site) = PASensor_AREA.MV(5).MinuendRegisterData(Site) - PASensor_AREA.MV(5).SubtrahendRegisterData(Site)
                        If PASensor_AREA.MV(0).EfuseData(Site) < -15 Or PASensor_AREA.MV(0).EfuseData(Site) > 15 Or PASensor_AREA.MV(1).EfuseData(Site) < -15 Or PASensor_AREA.MV(1).EfuseData(Site) > 15 Or PASensor_AREA.MV(4).EfuseData(Site) < -15 Or PASensor_AREA.MV(4).EfuseData(Site) > 15 Or PASensor_AREA.MV(5).EfuseData(Site) < -15 Or PASensor_AREA.MV(5).EfuseData(Site) > 15 Then
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.MV(0).EfuseData, TName:="Delta_FCM_LVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.MV(1).EfuseData, TName:="Delta_FCM_SVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.MV(4).EfuseData, TName:="Delta_NPU_LVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.MV(5).EfuseData, TName:="Delta_NPU_SVT"
                        Else
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("FCM_LVT", 0, Site, PASensor_AREA.MV(0).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("FCM_SVT", 0, Site, PASensor_AREA.MV(1).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("NPU_LVT", 0, Site, PASensor_AREA.MV(4).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("NPU_SVT", 0, Site, PASensor_AREA.MV(5).EfuseData(Site), 0, 0)
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("FCM_LVT")).gDataValue(Site) = CDbl(PASensor_AREA.MV(0).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("FCM_SVT")).gDataValue(Site) = CDbl(PASensor_AREA.MV(1).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("NPU_LVT")).gDataValue(Site) = CDbl(PASensor_AREA.MV(4).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("NPU_SVT")).gDataValue(Site) = CDbl(PASensor_AREA.MV(5).EfuseData(Site))
                        End If
                    Next Site
                Else
                    TheExec.Datalog.WriteComment "ERROR, don't support the size of Minuend or Subtrahend "
                    Exit Function
                End If
            ElseIf i_TestCondition_str(i) = "0P8" Then
                If i_EFUSEMinuend_lng + 1 = 9 And i_EFUSESubtrahend_lng + 1 = 9 Then
                    For Each Site In TheExec.Sites
                        PASensor_AREA.HV(6).EfuseData(Site) = PASensor_AREA.HV(6).MinuendRegisterData(Site) - PASensor_AREA.HV(6).SubtrahendRegisterData(Site)
                        PASensor_AREA.HV(7).EfuseData(Site) = PASensor_AREA.HV(7).MinuendRegisterData(Site) - PASensor_AREA.HV(7).SubtrahendRegisterData(Site)
                        PASensor_AREA.HV(8).EfuseData(Site) = PASensor_AREA.HV(8).MinuendRegisterData(Site) - PASensor_AREA.HV(8).SubtrahendRegisterData(Site)
                        If PASensor_AREA.HV(6).EfuseData(Site) < -15 Or PASensor_AREA.HV(6).EfuseData(Site) > 15 Or PASensor_AREA.HV(7).EfuseData(Site) < -15 Or PASensor_AREA.HV(7).EfuseData(Site) > 15 Or PASensor_AREA.HV(8).EfuseData(Site) < -15 Or PASensor_AREA.HV(8).EfuseData(Site) > 15 Then
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.HV(6).EfuseData, TName:="Delta_MODEM_ULVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.HV(7).EfuseData, TName:="Delta_MODEM5G_ULVT"
                            TheExec.Flow.TestLimit ResultVal:=PASensor_AREA.HV(8).EfuseData, TName:="Delta_BBP5G_ULVT"
                        Else
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("MODEM_ULVT", 0, Site, PASensor_AREA.HV(6).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("MODEM5G_ULVT", 0, Site, PASensor_AREA.HV(7).EfuseData(Site), 0, 0)
                            Call hiefuse.MULT_EFUSE_ITEM_UPDATE("BBP5G_ULVT", 0, Site, PASensor_AREA.HV(8).EfuseData(Site), 0, 0)
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("MODEM_ULVT")).gDataValue(Site) = CDbl(PASensor_AREA.HV(6).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("MODEM5G_ULVT")).gDataValue(Site) = CDbl(PASensor_AREA.HV(7).EfuseData(Site))
                            Data_Trim_AREA.PreTrim(PreTrimHashTable.Value("BBP5G_ULVT")).gDataValue(Site) = CDbl(PASensor_AREA.HV(8).EfuseData(Site))
                        End If
                    Next Site
                Else
                    TheExec.Datalog.WriteComment "ERROR, don't support the size of Minuend or Subtrahend "
                    Exit Function
                End If
            Else
                TheExec.Datalog.WriteComment " Please check Test Condition"
                Exit Function
            End If
        Next i
    End If

    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Sub Init_PASensor_Array()
ReDim g_PASensor_RegName_str(59)
ReDim g_PASensor_ValidRegName_str(14)

g_PASensor_RegName_str(0) = "FCM_WRAP0_pasensor_nbti_a_data"
g_PASensor_RegName_str(1) = "FCM_WRAP0_pasensor_nbti_o_data"
g_PASensor_RegName_str(2) = "FCM_WRAP0_pasensor_hci_a_data"
g_PASensor_RegName_str(3) = "FCM_WRAP0_pasensor_hci_o_data"
g_PASensor_RegName_str(4) = "FCM_WRAP1_pasensor_nbti_a_data"
g_PASensor_RegName_str(5) = "FCM_WRAP1_pasensor_nbti_o_data"
g_PASensor_RegName_str(6) = "FCM_WRAP1_pasensor_hci_a_data"
g_PASensor_RegName_str(7) = "FCM_WRAP1_pasensor_hci_o_data"
g_PASensor_RegName_str(8) = "FCM_WRAP2_pasensor_nbti_a_data"
g_PASensor_RegName_str(9) = "FCM_WRAP2_pasensor_nbti_o_data"
g_PASensor_RegName_str(10) = "FCM_WRAP2_pasensor_hci_a_data"
g_PASensor_RegName_str(11) = "FCM_WRAP2_pasensor_hci_o_data"
g_PASensor_RegName_str(12) = "DMC_TOP_0_pasensor_nbti_a_data"
g_PASensor_RegName_str(13) = "DMC_TOP_0_pasensor_nbti_o_data"
g_PASensor_RegName_str(14) = "DMC_TOP_0_pasensor_hci_a_data"
g_PASensor_RegName_str(15) = "DMC_TOP_0_pasensor_hci_o_data"
g_PASensor_RegName_str(16) = "DMC_TOP_0_1_pasensor_nbti_a_data"
g_PASensor_RegName_str(17) = "DMC_TOP_0_1_pasensor_nbti_o_data"
g_PASensor_RegName_str(18) = "DMC_TOP_0_1_pasensor_hci_a_data"
g_PASensor_RegName_str(19) = "DMC_TOP_0_1_pasensor_hci_o_data"
g_PASensor_RegName_str(20) = "NPU_WRAP0_0_pasensor_nbti_a_data"
g_PASensor_RegName_str(21) = "NPU_WRAP0_0_pasensor_nbti_o_data"
g_PASensor_RegName_str(22) = "NPU_WRAP0_0_pasensor_hci_a_data"
g_PASensor_RegName_str(23) = "NPU_WRAP0_0_pasensor_hci_o_data"
g_PASensor_RegName_str(24) = "NPU_WRAP1_0_pasensor_nbti_a_data"
g_PASensor_RegName_str(25) = "NPU_WRAP1_0_pasensor_nbti_o_data"
g_PASensor_RegName_str(26) = "NPU_WRAP1_0_pasensor_hci_a_data"
g_PASensor_RegName_str(27) = "NPU_WRAP1_0_pasensor_hci_o_data"
g_PASensor_RegName_str(28) = "NPU_CPU_WRAP0_pasensor_nbti_a_data"
g_PASensor_RegName_str(29) = "NPU_CPU_WRAP0_pasensor_nbti_o_data"
g_PASensor_RegName_str(30) = "NPU_CPU_WRAP0_pasensor_hci_a_data"
g_PASensor_RegName_str(31) = "NPU_CPU_WRAP0_pasensor_hci_o_data"
g_PASensor_RegName_str(32) = "NPU_CPU_WRAP1_pasensor_nbti_a_data"
g_PASensor_RegName_str(33) = "NPU_CPU_WRAP1_pasensor_nbti_o_data"
g_PASensor_RegName_str(34) = "NPU_CPU_WRAP1_pasensor_hci_a_data"
g_PASensor_RegName_str(35) = "NPU_CPU_WRAP1_pasensor_hci_o_data"
g_PASensor_RegName_str(36) = "G3D_LVT_pasensor_nbti_a_data"
g_PASensor_RegName_str(37) = "G3D_LVT_pasensor_nbti_o_data"
g_PASensor_RegName_str(38) = "G3D_LVT_pasensor_hci_a_data"
g_PASensor_RegName_str(39) = "G3D_LVT_pasensor_hci_o_data"
g_PASensor_RegName_str(40) = "G3D_SVT_pasensor_nbti_a_data"
g_PASensor_RegName_str(41) = "G3D_SVT_pasensor_nbti_o_data"
g_PASensor_RegName_str(42) = "G3D_SVT_pasensor_hci_a_data"
g_PASensor_RegName_str(43) = "G3D_SVT_pasensor_hci_o_data"
g_PASensor_RegName_str(44) = "G3D_ULVT_pasensor_nbti_a_data"
g_PASensor_RegName_str(45) = "G3D_ULVT_pasensor_nbti_o_data"
g_PASensor_RegName_str(46) = "G3D_ULVT_pasensor_hci_a_data"
g_PASensor_RegName_str(47) = "G3D_ULVT_pasensor_hci_o_data"
g_PASensor_RegName_str(48) = "bbp5g_ULVT_BBP5G_pasensor_nbti_a_data"
g_PASensor_RegName_str(49) = "bbp5g_ULVT_BBP5G_pasensor_nbti_o_data"
g_PASensor_RegName_str(50) = "bbp5g_ULVT_BBP5G_pasensor_hci_a_data"
g_PASensor_RegName_str(51) = "bbp5g_ULVT_BBP5G_pasensor_hci_o_data"
g_PASensor_RegName_str(52) = "MODEM5G_ULVT_MDM5G_pasensor_nbti_a_data"
g_PASensor_RegName_str(53) = "MODEM5G_ULVT_MDM5G_pasensor_nbti_o_data"
g_PASensor_RegName_str(54) = "MODEM5G_ULVT_MDM5G_pasensor_hci_a_data"
g_PASensor_RegName_str(55) = "MODEM5G_ULVT_MDM5G_pasensor_hci_o_data"
g_PASensor_RegName_str(56) = "MODEM_MDM4G_pasensor_nbti_a_data"
g_PASensor_RegName_str(57) = "MODEM_MDM4G_pasensor_nbti_o_data"
g_PASensor_RegName_str(58) = "MODEM_MDM4G_pasensor_hci_a_data"
g_PASensor_RegName_str(59) = "MODEM_MDM4G_pasensor_hci_o_data"
g_PASensor_ValidRegName_str(0) = "FCM_WRAP0_pasensor_valid"
g_PASensor_ValidRegName_str(1) = "FCM_WRAP1_pasensor_valid"
g_PASensor_ValidRegName_str(2) = "FCM_WRAP2_pasensor_valid"
g_PASensor_ValidRegName_str(3) = "DMC_ULVT_0_pasensor_valid"
g_PASensor_ValidRegName_str(4) = "DMC_ULVT_0_1_pasensor_valid"
g_PASensor_ValidRegName_str(5) = "NPU_WRAP0_0_pasensor_valid"
g_PASensor_ValidRegName_str(6) = "NPU_WRAP1_0_pasensor_valid"
g_PASensor_ValidRegName_str(7) = "NPU_CPU_WRAP0_pasensor_valid"
g_PASensor_ValidRegName_str(8) = "NPU_CPU_WRAP1_pasensor_valid"
g_PASensor_ValidRegName_str(9) = "G3D_LVT_pasensor_valid"
g_PASensor_ValidRegName_str(10) = "G3D_SVT_pasensor_valid"
g_PASensor_ValidRegName_str(11) = "G3D_ULVT_pasensor_valid"
g_PASensor_ValidRegName_str(12) = "bbp5g_ULVT_BBP5G_pasensor_valid"
g_PASensor_ValidRegName_str(13) = "MODEM5GULVT_MDM5G_pasensor_valid"
g_PASensor_ValidRegName_str(14) = "MODEM_ULVT_MDM4G_pasensor_valid"

End Sub



