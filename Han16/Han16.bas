Option Explicit

'============================================================
' Han16 / Han16-NIR batch computation (Excel VBA)
'
' Inputs
'   - Rrs665 in column H
'   - Rrs753 in column I
'   - Coefficients in fixed cells (see constants below)
'
' Outputs
'   - Han16      -> column J
'   - Han16_NIR  -> column K
'   - Han_Type   -> column L  ("L", "HR", "NIR")
'
' Regime rule (based on Rrs665)
'   - L   : Rrs665 <= 0.03
'   - HR  : 0.03 < Rrs665 < 0.04  (log10-weighted interpolation)
'   - NIR : Rrs665 >= 0.04
'============================================================
Public Sub Han16Mtd()

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    ' ----- Column mapping -----
    Const COL_RRS665 As String = "H"
    Const COL_RRS753 As String = "I"
    Const COL_OUT_HAN16 As String = "J"
    Const COL_OUT_HAN16NIR As String = "K"
    Const COL_OUT_TYPE As String = "L"

    ' ----- Coefficient cells -----
    Const CELL_A_LOW  As String = "B2"
    Const CELL_C_LOW  As String = "D2"
    Const CELL_A_HIGH As String = "E2"
    Const CELL_C_HIGH As String = "G2"
    Const CELL_A_NIR  As String = "E3"
    Const CELL_C_NIR  As String = "G3"

    ' ----- Constants -----
    Const PI As Double = 3.14159265358979
    Const THRESH_LOW As Double = 0.03
    Const THRESH_HIGH As Double = 0.04

    ' ----- Coefficients -----
    Dim A_low As Double, C_low As Double
    Dim A_high As Double, C_high As Double
    Dim A_NIR As Double, C_NIR As Double

    ' ----- Per-row variables -----
    Dim Rrs665 As Double, Rrs753 As Double
    Dim WL As Double, WH As Double
    Dim SPM_L As Double, SPM_NIR As Double, SPM_HR As Double, SPM_HNIR As Double
    Dim mark As String

    Set ws = ThisWorkbook.Sheets(1)

    ' Read coefficients
    A_low = Coeff(ws, CELL_A_LOW)
    C_low = Coeff(ws, CELL_C_LOW)
    A_high = Coeff(ws, CELL_A_HIGH)
    C_high = Coeff(ws, CELL_C_HIGH)
    A_NIR = Coeff(ws, CELL_A_NIR)
    C_NIR = Coeff(ws, CELL_C_NIR)

    ' Headers
    ws.Cells(1, COL_OUT_HAN16).Value = "Han16"
    ws.Cells(1, COL_OUT_HAN16NIR).Value = "Han16_NIR"
    ws.Cells(1, COL_OUT_TYPE).Value = "Han_Type"

    ' Last data row (based on Rrs665 column)
    lastRow = ws.Cells(ws.Rows.Count, COL_RRS665).End(xlUp).Row

    For r = 2 To lastRow

        ' Defaults
        ws.Cells(r, COL_OUT_HAN16).Value = "NaN"
        ws.Cells(r, COL_OUT_HAN16NIR).Value = "NaN"
        ws.Cells(r, COL_OUT_TYPE).Value = "NaN"

        ' Read inputs
        If Not IsNumeric(ws.Cells(r, COL_RRS665).Value) Then GoTo NextRow
        Rrs665 = ws.Cells(r, COL_RRS665).Value
        If Rrs665 <= 0 Then GoTo NextRow

        If IsNumeric(ws.Cells(r, COL_RRS753).Value) Then
            Rrs753 = ws.Cells(r, COL_RRS753).Value
        Else
            Rrs753 = 0
        End If

        ' Regime + weights
        mark = GetHanMarkAndWeights(Rrs665, THRESH_LOW, THRESH_HIGH, WL, WH)

        ' Candidates
        SPM_L = SPM(A_low, Rrs665, C_low, PI)
        SPM_NIR = SPM(A_high, Rrs665, C_high, PI)
        SPM_HR = (WL * SPM_L + WH * SPM_NIR) / (WL + WH)
        SPM_HNIR = SPM(A_NIR, Rrs753, C_NIR, PI)

        ' Assign outputs
        Select Case mark
            Case "L"
                ws.Cells(r, COL_OUT_HAN16).Value = SPM_L
                ws.Cells(r, COL_OUT_HAN16NIR).Value = SPM_L
            Case "HR"
                ws.Cells(r, COL_OUT_HAN16).Value = SPM_HR
                ws.Cells(r, COL_OUT_HAN16NIR).Value = SPM_HR
            Case "NIR"
                ws.Cells(r, COL_OUT_HAN16).Value = SPM_NIR
                ws.Cells(r, COL_OUT_HAN16NIR).Value = SPM_HNIR
        End Select

        ws.Cells(r, COL_OUT_TYPE).Value = mark

NextRow:
    Next r

End Sub

' Returns the coefficient value from a fixed cell (non-numeric -> 0).
Private Function Coeff(ByVal ws As Worksheet, ByVal addr As String) As Double
    If IsNumeric(ws.Range(addr).Value) Then
        Coeff = CDbl(ws.Range(addr).Value)
    Else
        Coeff = 0
    End If
End Function

' Core SPM formula: A * Rrs * pi / (1 - Rrs*pi/C).
Private Function SPM(ByVal A As Double, ByVal Rrs As Double, ByVal C As Double, ByVal PI As Double) As Double
    SPM = A * Rrs * PI / (1 - Rrs * PI / C)
End Function

' Returns regime label and sets WL/WH.
Private Function GetHanMarkAndWeights( _
    ByVal Rrs665 As Double, _
    ByVal thrLow As Double, _
    ByVal thrHigh As Double, _
    ByRef WL As Double, _
    ByRef WH As Double) As String

    If Rrs665 <= thrLow Then
        WL = 1: WH = 0
        GetHanMarkAndWeights = "L"
    ElseIf Rrs665 >= thrHigh Then
        WL = 0: WH = 1
        GetHanMarkAndWeights = "NIR"
    Else
        WL = WorksheetFunction.Log10(thrHigh) - WorksheetFunction.Log10(Rrs665)
        WH = WorksheetFunction.Log10(Rrs665) - WorksheetFunction.Log10(thrLow)
        GetHanMarkAndWeights = "HR"
    End If

End Function
