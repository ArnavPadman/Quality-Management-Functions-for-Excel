Attribute VB_Name = "Module1"
Option Explicit

Private Const FUNCTION_CATEGORY As String = "Quality & Six Sigma"


' =========================================================
' STANDARD ERROR
' =========================================================

Public Function QMS_STANDARD_ERROR( _
    ByVal data As Range) As Variant
Attribute QMS_STANDARD_ERROR.VB_Description = "Returns the standard error of the sample mean."
Attribute QMS_STANDARD_ERROR.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim n As Long

    n = Application.WorksheetFunction.Count(data)

    If n < 2 Then
        QMS_STANDARD_ERROR = CVErr(xlErrDiv0)
        Exit Function
    End If

    QMS_STANDARD_ERROR = _
        Application.WorksheetFunction.StDev_S(data) / Sqr(n)

End Function


' =========================================================
' TAGUCHI LOSS — INDIVIDUAL MEASUREMENT
' =========================================================

Public Function QMS_TAGUCHI_LOSS( _
    ByVal actual As Double, _
    ByVal target As Double, _
    ByVal loss_at_tolerance As Double, _
    ByVal tolerance As Double) As Variant
Attribute QMS_TAGUCHI_LOSS.VB_Description = "Returns Taguchi loss for one individual measurement."
Attribute QMS_TAGUCHI_LOSS.VB_ProcData.VB_Invoke_Func = " \n20"

    If tolerance <= 0 Or loss_at_tolerance < 0 Then
        QMS_TAGUCHI_LOSS = CVErr(xlErrNum)
        Exit Function
    End If

    QMS_TAGUCHI_LOSS = _
        (loss_at_tolerance / tolerance ^ 2) * _
        (actual - target) ^ 2

End Function


' =========================================================
' TAGUCHI LOSS — PROCESS SAMPLE
' =========================================================

Public Function QMS_TAGUCHI_PROCESS_LOSS( _
    ByVal data As Range, _
    ByVal target As Double, _
    ByVal loss_at_tolerance As Double, _
    ByVal tolerance As Double) As Variant
Attribute QMS_TAGUCHI_PROCESS_LOSS.VB_Description = "Returns the estimated average Taguchi loss per unit for a process sample."
Attribute QMS_TAGUCHI_PROCESS_LOSS.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim n As Long
    Dim sampleMean As Double
    Dim sampleSD As Double
    Dim lossConstant As Double

    n = Application.WorksheetFunction.Count(data)

    If n < 2 Then
        QMS_TAGUCHI_PROCESS_LOSS = CVErr(xlErrDiv0)
        Exit Function
    End If

    If tolerance <= 0 Or loss_at_tolerance < 0 Then
        QMS_TAGUCHI_PROCESS_LOSS = CVErr(xlErrNum)
        Exit Function
    End If

    sampleMean = Application.WorksheetFunction.Average(data)
    sampleSD = Application.WorksheetFunction.StDev_S(data)

    lossConstant = loss_at_tolerance / tolerance ^ 2

    QMS_TAGUCHI_PROCESS_LOSS = _
        lossConstant * _
        (sampleSD ^ 2 + (sampleMean - target) ^ 2)

End Function


' =========================================================
' THROUGHPUT YIELD
' =========================================================

Public Function QMS_THROUGHPUT_YIELD( _
    ByVal good_units As Double, _
    ByVal total_units As Double) As Variant
Attribute QMS_THROUGHPUT_YIELD.VB_Description = "Returns the yield for one process step."
Attribute QMS_THROUGHPUT_YIELD.VB_ProcData.VB_Invoke_Func = " \n20"

    If total_units <= 0 Then
        QMS_THROUGHPUT_YIELD = CVErr(xlErrNum)
        Exit Function
    End If

    If good_units < 0 Or good_units > total_units Then
        QMS_THROUGHPUT_YIELD = CVErr(xlErrNum)
        Exit Function
    End If

    QMS_THROUGHPUT_YIELD = good_units / total_units

End Function


' =========================================================
' ROLLED THROUGHPUT YIELD
' =========================================================

Public Function QMS_ROLLED_YIELD( _
    ByVal stage_yields As Range) As Variant
Attribute QMS_ROLLED_YIELD.VB_Description = "Returns rolled throughput yield by multiplying all stage yields."
Attribute QMS_ROLLED_YIELD.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim cell As Range
    Dim cellValue As Variant
    Dim result As Double
    Dim numericCount As Long

    result = 1

    For Each cell In stage_yields.Cells

        cellValue = cell.Value2

        If IsError(cellValue) Then
            QMS_ROLLED_YIELD = CVErr(xlErrValue)
            Exit Function
        End If

        If Len(cellValue) > 0 Then

            If Not IsNumeric(cellValue) Then
                QMS_ROLLED_YIELD = CVErr(xlErrValue)
                Exit Function
            End If

            If CDbl(cellValue) < 0 Or CDbl(cellValue) > 1 Then
                QMS_ROLLED_YIELD = CVErr(xlErrNum)
                Exit Function
            End If

            result = result * CDbl(cellValue)
            numericCount = numericCount + 1

        End If

    Next cell

    If numericCount = 0 Then
        QMS_ROLLED_YIELD = CVErr(xlErrValue)
    Else
        QMS_ROLLED_YIELD = result
    End If

End Function


' =========================================================
' DEFECTS PER MILLION OPPORTUNITIES
' =========================================================

Public Function QMS_DPMO( _
    ByVal defects As Double, _
    ByVal units As Double, _
    ByVal opportunities_per_unit As Double) As Variant
Attribute QMS_DPMO.VB_Description = "Returns defects per million opportunities."
Attribute QMS_DPMO.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim totalOpportunities As Double

    If defects < 0 Then
        QMS_DPMO = CVErr(xlErrNum)
        Exit Function
    End If

    If units <= 0 Or opportunities_per_unit <= 0 Then
        QMS_DPMO = CVErr(xlErrNum)
        Exit Function
    End If

    totalOpportunities = units * opportunities_per_unit

    If defects > totalOpportunities Then
        QMS_DPMO = CVErr(xlErrNum)
        Exit Function
    End If

    QMS_DPMO = _
        defects / totalOpportunities * 1000000#

End Function


' =========================================================
' SIGMA LEVEL
' =========================================================

Public Function QMS_SIGMA_LEVEL( _
    ByVal dpmo As Double, _
    Optional ByVal sigma_shift As Double = 1.5) As Variant
Attribute QMS_SIGMA_LEVEL.VB_Description = "Converts DPMO to an estimated sigma level."
Attribute QMS_SIGMA_LEVEL.VB_ProcData.VB_Invoke_Func = " \n20"

    If dpmo <= 0 Or dpmo >= 1000000# Then
        QMS_SIGMA_LEVEL = CVErr(xlErrNum)
        Exit Function
    End If

    QMS_SIGMA_LEVEL = _
        Application.WorksheetFunction.Norm_S_Inv( _
            1 - dpmo / 1000000#) + sigma_shift

End Function


' =========================================================
' REGISTER FUNCTION DESCRIPTIONS
' =========================================================

Public Sub RegisterQualityFunctions()

    On Error Resume Next

    RegisterFunctionHelp _
        "QMS_STANDARD_ERROR", _
        "Returns the standard error of the sample mean.", _
        "data (range): Sample observations. At least two numbers are required."

    RegisterFunctionHelp _
        "QMS_TAGUCHI_LOSS", _
        "Returns Taguchi loss for one individual measurement.", _
        "actual (number): The measured value.", _
        "target (number): The desired target value.", _
        "loss_at_tolerance (number): Loss at the tolerance limit.", _
        "tolerance (number): Allowed deviation from target."

    RegisterFunctionHelp _
        "QMS_TAGUCHI_PROCESS_LOSS", _
        "Returns the estimated average Taguchi loss per unit for a process sample.", _
        "data (range): Sample measurements. At least two numbers are required.", _
        "target (number): The desired target value.", _
        "loss_at_tolerance (number): Loss at the tolerance limit.", _
        "tolerance (number): Allowed deviation from target."

    RegisterFunctionHelp _
        "QMS_THROUGHPUT_YIELD", _
        "Returns the yield for one process step.", _
        "good_units (number): Units that passed the process step.", _
        "total_units (number): Total units entering the process step."

    RegisterFunctionHelp _
        "QMS_ROLLED_YIELD", _
        "Returns rolled throughput yield by multiplying all stage yields.", _
        "stage_yields (range): Stage yields entered as decimals from 0 to 1."

    RegisterFunctionHelp _
        "QMS_DPMO", _
        "Returns defects per million opportunities.", _
        "defects (number): Total defects observed.", _
        "units (number): Number of units inspected.", _
        "opportunities_per_unit (number): Defect opportunities per unit."

    RegisterFunctionHelp _
        "QMS_SIGMA_LEVEL", _
        "Converts DPMO to an estimated sigma level.", _
        "dpmo (number): Defects per million opportunities.", _
        "sigma_shift (optional number): Defaults to 1.5. Use 0 for no shift."

    On Error GoTo 0

End Sub


' =========================================================
' INTERNAL HELP REGISTRATION
' =========================================================

Private Sub RegisterFunctionHelp( _
    ByVal functionName As String, _
    ByVal functionDescription As String, _
    ParamArray argumentHelp() As Variant)

    Application.MacroOptions _
        Macro:=functionName, _
        Description:=functionDescription, _
        Category:=FUNCTION_CATEGORY, _
        ArgumentDescriptions:=argumentHelp

End Sub


' =========================================================
' RUN WHEN THE ADD-IN OPENS
' =========================================================

Public Sub Auto_Open()

    RegisterQualityFunctions

End Sub

