Attribute VB_Name = "VBT_A_AVS"
Option Explicit

Public Type STR_HISI_AVS_TAB
    IntanceName As String
    Formula As String
    Coef_a As Double
    Coef_b As Double
    Coef_c As Double
    Coef_d As Double
    LowVoltage As Double
    HighVoltage As Double
    HPMTest As String
    powerPin As String
End Type

Public aSTR_HISI_AVS_TAB(50) As STR_HISI_AVS_TAB

Public gHPM_ULVT__enyo_u2_hpm_pc_org_MV         As New SiteDouble
Public gHPM_LVT__enyo_midcore_u2_hpm_pc_org_MV  As New SiteDouble
Public gHPM_LVT__ananke_u2_hpm_pc_org_MV        As New SiteDouble
Public gHPM_LVT__FCM_u2_hpm_pc_org_MV           As New SiteDouble
Public gHPM_LVT__G3D_u2_hpm_pc_org_MV           As New SiteDouble
Public gHPM_LVT__NPU_WRAP_u2_hpm_pc_org_MV      As New SiteDouble
Public gHPM_LVT__PERI_u1_hpm_pc_org_MV          As New SiteDouble


Public Function AVS_TABLE_INI() As Long
    Dim RW As Variant
    Dim index As Long
    Dim CurrentSheetName As String
    CurrentSheetName = ActiveSheet.Name
    Worksheets("AVS_TABLE").Activate
    
    For Each RW In Range("a2:a50")
        
        'to detect the empty cell in EFUSE Item Table
        If Cells(RW.row, 1) = "" Or Cells(RW.row, 2) = "" Or Cells(RW.row, 3) = "" Or Cells(RW.row, 4) = "" Or _
           Cells(RW.row, 5) = "" Or Cells(RW.row, 6) = "" Or Cells(RW.row, 6) = "" Or Cells(RW.row, 7) = "" Or _
           Cells(RW.row, 8) = "" Or Cells(RW.row, 9) = "" Or Cells(RW.row, 10) = "" Then
            Exit For
        Else
            aSTR_HISI_AVS_TAB(index).IntanceName = Cells(RW.row, 1)
            aSTR_HISI_AVS_TAB(index).Formula = Cells(RW.row, 2)
            aSTR_HISI_AVS_TAB(index).Coef_a = CDbl(Cells(RW.row, 3))
            aSTR_HISI_AVS_TAB(index).Coef_b = CDbl(Cells(RW.row, 4))
            aSTR_HISI_AVS_TAB(index).Coef_c = CDbl(Cells(RW.row, 5))
            aSTR_HISI_AVS_TAB(index).Coef_d = CDbl(Cells(RW.row, 6))
            aSTR_HISI_AVS_TAB(index).LowVoltage = CDbl(Cells(RW.row, 7))
            aSTR_HISI_AVS_TAB(index).HighVoltage = CDbl(Cells(RW.row, 8))
            aSTR_HISI_AVS_TAB(index).HPMTest = Cells(RW.row, 9)
            aSTR_HISI_AVS_TAB(index).powerPin = Cells(RW.row, 10)
        End If
        index = index + 1
    Next RW
    Worksheets(CurrentSheetName).Activate
    
    'Initial Gobal Variant
    gHPM_ULVT__enyo_u2_hpm_pc_org_MV = -1
    gHPM_LVT__enyo_midcore_u2_hpm_pc_org_MV = -1
    gHPM_LVT__ananke_u2_hpm_pc_org_MV = -1
    gHPM_LVT__FCM_u2_hpm_pc_org_MV = -1
    gHPM_LVT__G3D_u2_hpm_pc_org_MV = -1
    gHPM_LVT__NPU_WRAP_u2_hpm_pc_org_MV = -1
    gHPM_LVT__PERI_u1_hpm_pc_org_MV = -1
End Function

Public Function AVS_ApplyLevel(wait_time As Double, Debug_Mode As Boolean) As Long
On Error GoTo errHandler
    Dim index As Long
    Dim CurInsName As String
    Dim y_Val As New SiteDouble
    Dim x_Val As New SiteDouble
    Dim PinName As String
    Dim y_Val_Max As New SiteDouble
    Dim Site As Variant

    Dim FindInstHash As New clsHashTable
    Dim Key As Variant
    Dim keys() As String
    
    FindInstHash.CaseSensitiveKeys = True
    
    'Find all Instance Index in AVS Table
    CurInsName = TheExec.DataManager.InstanceName
    
    'Bypass non-TRANS instances
    If InStr(CurInsName, "FC_TRANS") = 0 Then Exit Function
    
    For index = 0 To UBound(aSTR_HISI_AVS_TAB)
        If InStr(aSTR_HISI_AVS_TAB(index).IntanceName, CurInsName) > 0 Then
            If FindInstHash.KeyExists(CStr(index)) = False Then
                FindInstHash.Add CStr(index), CurInsName
            Else
                TheExec.Datalog.WriteComment "Error: Found duplicate Index in AVS table!!!"
                Exit Function
            End If
        
        End If
    Next index
    
    If FindInstHash.count >= 1 Then
        'Do Calculation and find the maximum
        keys = FindInstHash.keys
        For Each Key In keys
            index = CLng(Key)
            'Select HPM Test Value
            Select Case aSTR_HISI_AVS_TAB(index).HPMTest
                Case "HPM_ULVT__enyo_u2_hpm_pc_org_MV"
                    x_Val = gHPM_ULVT__enyo_u2_hpm_pc_org_MV
                Case "HPM_LVT__enyo_midcore_u2_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__enyo_midcore_u2_hpm_pc_org_MV
                Case "HPM_LVT__ananke_u2_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__ananke_u2_hpm_pc_org_MV
                Case "HPM_LVT__FCM_u2_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__FCM_u2_hpm_pc_org_MV
                Case "HPM_LVT__G3D_u2_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__G3D_u2_hpm_pc_org_MV
                Case "HPM_LVT__NPU_WRAP_u2_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__NPU_WRAP_u2_hpm_pc_org_MV
                Case "HPM_LVT__PERI_u1_hpm_pc_org_MV"
                    x_Val = gHPM_LVT__PERI_u1_hpm_pc_org_MV
                Case Else
                    TheExec.Datalog.WriteComment "Error: Get wrong HPM Test Name From AVS table!!!"
                    TheExec.Flow.TestLimit 0, 1, 1
            End Select

            'Get the power pin name
            PinName = aSTR_HISI_AVS_TAB(index).powerPin
            
            'Do the calcution of voltage serially
            For Each Site In TheExec.Sites
                If x_Val < 0 Then
                    TheExec.Datalog.WriteComment "Error: HPM Test does not send correct value to AVS test in Site" & CStr(Site) & "!!!"
                    TheExec.Flow.TestLimit 0, 1, 1
                End If
                
                'Calculate the Y Value
                'Pwor Value Formula: y = a - b * x + d * c
                y_Val = aSTR_HISI_AVS_TAB(index).Coef_a - aSTR_HISI_AVS_TAB(index).Coef_b * x_Val + aSTR_HISI_AVS_TAB(index).Coef_d * aSTR_HISI_AVS_TAB(index).Coef_c
                
                'Set the boundary
                If y_Val > aSTR_HISI_AVS_TAB(index).HighVoltage Then
                    y_Val = aSTR_HISI_AVS_TAB(index).HighVoltage
                ElseIf y_Val < aSTR_HISI_AVS_TAB(index).LowVoltage Then
                    y_Val = aSTR_HISI_AVS_TAB(index).LowVoltage
                End If
                
                'Set the unit
                y_Val = y_Val * MV
                
                'Find the maximum of y_val
                If y_Val > y_Val_Max Then
                    y_Val_Max = y_Val
                End If
                If True Then
                    TheExec.Flow.TestLimit x_Val, TName:="X_Val:" ' & CStr(Site)
                    TheExec.Flow.TestLimit y_Val, TName:="Y_Val:"  ' & CStr(Site)
                End If
            Next Site
            If Debug_Mode Then
                TheExec.Datalog.WriteComment CurInsName & "'s AVS Formula is: " & aSTR_HISI_AVS_TAB(index).Formula
                TheExec.Datalog.WriteComment CurInsName & "'s AVS HPM Test is: " & aSTR_HISI_AVS_TAB(index).HPMTest
                TheExec.Datalog.WriteComment CurInsName & "'s High Voltage is: " & aSTR_HISI_AVS_TAB(index).HighVoltage & "mv"
                TheExec.Datalog.WriteComment CurInsName & "'s Low Voltage is: " & aSTR_HISI_AVS_TAB(index).LowVoltage & "mv"
            End If
        Next Key
        
        'Apply Level to Power Pin
        If TheExec.TesterMode = testModeOnline Then
            TheHdw.DCVS.Pins(PinName).voltage.Main.SiteAwareValue = y_Val_Max
            TheHdw.Wait wait_time
        End If

        If Debug_Mode Then
            For Each Site In TheExec.Sites
                TheExec.Flow.TestLimit TheHdw.DCVS.Pins(PinName).voltage.Main.Value, TName:=CurInsName & "_Readback", PinName:=PinName
            Next Site
        End If
    End If
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

