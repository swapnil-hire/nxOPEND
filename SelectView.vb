Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Annotations

Module DrawingMasterComparison

    Dim theSession As Session = Session.GetSession()
    Dim theUFSession As UFSession = UFSession.GetUFSession()
    Dim theUI As UI = UI.GetUI()
    Dim workPart As Part = theSession.Parts.Work

    '=========================================================================
    ' MASTER DATA STRUCTURE - Define your Design Master values here
    ' In production, load this from XML/Excel/Teamcenter instead
    '=========================================================================
    Structure MasterDimension
        Dim Name        As String
        Dim NominalValue As Double
        Dim UpperTol    As Double
        Dim LowerTol    As Double
        Dim DimType     As String   ' LINEAR, DIAMETER, RADIUS, ANGULAR
        Dim IsMandatory As Boolean
    End Structure

    Structure CompareResult
        Dim DimName       As String
        Dim MasterValue   As Double
        Dim ActualValue   As Double
        Dim UpperTol      As Double
        Dim LowerTol      As Double
        Dim Status        As String  ' PASS / FAIL / MISSING / EXTRA
        Dim Remarks       As String
    End Structure

    '=========================================================================
    ' MAIN ENTRY POINT
    '=========================================================================
    Sub Main()

        If workPart Is Nothing Then
            MsgBox("No active part found.")
            Return
        End If

        Dim drawingSheet As Drawings.DrawingSheet = workPart.DrawingSheets.CurrentDrawingSheet
        If drawingSheet Is Nothing Then
            MsgBox("No active drawing sheet. Please open a drawing.")
            Return
        End If

        ' Step 1: Select a view interactively
        Dim selectedView As Drawings.DraftingView = SelectViewInteractive()
        If selectedView Is Nothing Then
            MsgBox("No view selected. Exiting.")
            Return
        End If

        MsgBox("Selected View: " & selectedView.Name & vbNewLine & "Extracting data...")

        ' Step 2: Extract all dimensions from the selected view
        Dim extractedDims As List(Of MasterDimension) = ExtractDimensionsFromView(selectedView)

        ' Step 3: Extract PMI annotations from the 3D model (design master source)
        Dim masterDims As List(Of MasterDimension) = ExtractPMIFromModel()

        ' Step 4: Compare drawing dimensions vs master
        Dim results As List(Of CompareResult) = CompareDimensions(extractedDims, masterDims)

        ' Step 5: Highlight discrepancies in NX view
        HighlightDiscrepancies(selectedView, results)

        ' Step 6: Generate comparison report
        GenerateReport(selectedView, results)

    End Sub

    '=========================================================================
    ' STEP 1 - Interactive View Selection
    '=========================================================================
    Function SelectViewInteractive() As Drawings.DraftingView

        Dim maskTriples(0) As Selection.MaskTriple
        maskTriples(0).Type    = UFConstants.UF_drawing_view_type
        maskTriples(0).Subtype = 0
        maskTriples(0).SolidBodySubtype = 0

        Dim selectedObjects() As TaggedObject = Nothing

        Dim response As Selection.Response = theUI.SelectionManager.SelectTaggedObjects(
            "Click on a Drawing View to Verify",
            "Select View",
            Selection.Scope.AnyInAssembly,
            Selection.Action.SelectOne,
            False,
            False,
            maskTriples,
            selectedObjects)

        If (response = Selection.Response.Ok OrElse
            response = Selection.Response.ObjectSelected) AndAlso
            selectedObjects IsNot Nothing AndAlso
            selectedObjects.Length > 0 Then

            Return TryCast(selectedObjects(0), Drawings.DraftingView)
        End If

        Return Nothing

    End Function

    '=========================================================================
    ' STEP 2 - Extract Dimensions from Drawing View
    '=========================================================================
    Function ExtractDimensionsFromView(dView As Drawings.DraftingView) As List(Of MasterDimension)

        Dim dimList As New List(Of MasterDimension)

        ' Get all objects visible in the view
        Dim viewTag As Tag = dView.Tag
        Dim numObjects As Integer = 0
        Dim objectTags() As Tag = Nothing

        theUFSession.View.AskVisibleObjects(viewTag, numObjects, objectTags)

        ' Iterate through all annotations in the part
        For Each ann As Annotation In workPart.Annotations

            ' Check if annotation belongs to this view
            If IsAnnotationInView(ann, dView) Then

                Dim md As New MasterDimension
                md.IsMandatory = True

                ' ---- Linear Dimension ----
                Dim linDim As LinearDimension = TryCast(ann, LinearDimension)
                If linDim IsNot Nothing Then
                    md.Name         = "LIN_" & linDim.Tag.ToString()
                    md.DimType      = "LINEAR"
                    md.NominalValue = GetDimensionValue(linDim)
                    GetToleranceValues(linDim, md.UpperTol, md.LowerTol)
                    dimList.Add(md)
                    Continue For
                End If

                ' ---- Diameter Dimension ----
                Dim diaDim As DiameterDimension = TryCast(ann, DiameterDimension)
                If diaDim IsNot Nothing Then
                    md.Name         = "DIA_" & diaDim.Tag.ToString()
                    md.DimType      = "DIAMETER"
                    md.NominalValue = GetDimensionValue(diaDim)
                    GetToleranceValues(diaDim, md.UpperTol, md.LowerTol)
                    dimList.Add(md)
                    Continue For
                End If

                ' ---- Radius Dimension ----
                Dim radDim As RadiusDimension = TryCast(ann, RadiusDimension)
                If radDim IsNot Nothing Then
                    md.Name         = "RAD_" & radDim.Tag.ToString()
                    md.DimType      = "RADIUS"
                    md.NominalValue = GetDimensionValue(radDim)
                    GetToleranceValues(radDim, md.UpperTol, md.LowerTol)
                    dimList.Add(md)
                    Continue For
                End If

                ' ---- Angular Dimension ----
                Dim angDim As AngularDimension = TryCast(ann, AngularDimension)
                If angDim IsNot Nothing Then
                    md.Name         = "ANG_" & angDim.Tag.ToString()
                    md.DimType      = "ANGULAR"
                    md.NominalValue = GetDimensionValue(angDim)
                    GetToleranceValues(angDim, md.UpperTol, md.LowerTol)
                    dimList.Add(md)
                    Continue For
                End If

                ' ---- GD&T Feature Control Frame ----
                Dim fcf As FeatureControlFrame = TryCast(ann, FeatureControlFrame)
                If fcf IsNot Nothing Then
                    md.Name         = "FCF_" & fcf.Tag.ToString()
                    md.DimType      = "GDT"
                    md.NominalValue = 0
                    md.UpperTol     = 0
                    md.LowerTol     = 0
                    dimList.Add(md)
                    Continue For
                End If

            End If

        Next

        Return dimList

    End Function

    '=========================================================================
    ' STEP 3 - Extract PMI from 3D Model (Design Master)
    '=========================================================================
    Function ExtractPMIFromModel() As List(Of MasterDimension)

        Dim masterList As New List(Of MasterDimension)

        ' Loop through all PMI objects in the part
        For Each pmiObj As NXObject In workPart.PMIManager.GetAllPMIObjects()

            Dim md As New MasterDimension
            md.IsMandatory = True

            ' ---- PMI Dimension ----
            Dim pmiDim As PMI.Dimension = TryCast(pmiObj, PMI.Dimension)
            If pmiDim IsNot Nothing Then
                md.Name         = "PMI_" & pmiDim.Name
                md.DimType      = "LINEAR"
                md.NominalValue = pmiDim.NominalValue
                md.UpperTol     = pmiDim.ToleranceUpper
                md.LowerTol     = pmiDim.ToleranceLower
                masterList.Add(md)
                Continue For
            End If

            ' ---- PMI GD&T ----
            Dim pmiGdt As PMI.FeatureControlFrame = TryCast(pmiObj, PMI.FeatureControlFrame)
            If pmiGdt IsNot Nothing Then
                md.Name         = "PMI_GDT_" & pmiGdt.Name
                md.DimType      = "GDT"
                md.NominalValue = 0
                masterList.Add(md)
                Continue For
            End If

            ' ---- PMI Surface Finish ----
            Dim pmiSF As PMI.SurfaceFinish = TryCast(pmiObj, PMI.SurfaceFinish)
            If pmiSF IsNot Nothing Then
                md.Name         = "PMI_SF_" & pmiSF.Name
                md.DimType      = "SURFACE_FINISH"
                md.NominalValue = 0
                masterList.Add(md)
                Continue For
            End If

        Next

        Return masterList

    End Function

    '=========================================================================
    ' STEP 4 - Compare Drawing Dimensions vs Master PMI
    '=========================================================================
    Function CompareDimensions(
        drawingDims As List(Of MasterDimension),
        masterDims  As List(Of MasterDimension)) As List(Of CompareResult)

        Dim results As New List(Of CompareResult)

        ' --- Check each master dim exists in drawing ---
        For Each master As MasterDimension In masterDims

            Dim result As New CompareResult
            result.DimName    = master.Name
            result.MasterValue = master.NominalValue
            result.UpperTol   = master.UpperTol
            result.LowerTol   = master.LowerTol

            ' Try to find matching dimension in drawing
            Dim matchFound As Boolean = False

            For Each drawn As MasterDimension In drawingDims

                If drawn.DimType = master.DimType Then

                    Dim valueDiff As Double = Math.Abs(drawn.NominalValue - master.NominalValue)

                    ' Match within 0.001 tolerance for floating point
                    If valueDiff < 0.001 Then
                        matchFound          = True
                        result.ActualValue  = drawn.NominalValue

                        ' Check if tolerance matches master
                        Dim tolDiff As Double = Math.Abs(drawn.UpperTol - master.UpperTol) +
                                                Math.Abs(drawn.LowerTol - master.LowerTol)

                        If tolDiff < 0.0001 Then
                            result.Status  = "PASS"
                            result.Remarks = "Dimension and tolerance match master."
                        Else
                            result.Status  = "FAIL"
                            result.Remarks = "Tolerance mismatch. Master: +" &
                                             master.UpperTol & "/-" & master.LowerTol &
                                             " | Drawing: +" &
                                             drawn.UpperTol & "/-" & drawn.LowerTol
                        End If

                        Exit For
                    End If
                End If

            Next

            If Not matchFound Then
                result.Status      = "MISSING"
                result.ActualValue = 0
                result.Remarks     = "Dimension exists in master but NOT found in drawing."
            End If

            results.Add(result)

        Next

        ' --- Check for EXTRA dims in drawing not in master ---
        For Each drawn As MasterDimension In drawingDims

            Dim foundInMaster As Boolean = False

            For Each master As MasterDimension In masterDims
                If drawn.DimType = master.DimType AndAlso
                   Math.Abs(drawn.NominalValue - master.NominalValue) < 0.001 Then
                    foundInMaster = True
                    Exit For
                End If
            Next

            If Not foundInMaster Then
                Dim extra As New CompareResult
                extra.DimName      = drawn.Name
                extra.MasterValue  = 0
                extra.ActualValue  = drawn.NominalValue
                extra.Status       = "EXTRA"
                extra.Remarks      = "Dimension found in drawing but NOT in design master."
                results.Add(extra)
            End If

        Next

        Return results

    End Function

    '=========================================================================
    ' STEP 5 - Highlight Discrepancies in NX View
    '=========================================================================
    Sub HighlightDiscrepancies(dView As Drawings.DraftingView,
                                results As List(Of CompareResult))

        ' Color codes: Red = FAIL/MISSING, Yellow = EXTRA, Green = PASS
        Dim RED    As Integer = 186  ' NX color index for Red
        Dim YELLOW As Integer = 6
        Dim GREEN  As Integer = 31

        For Each ann As Annotation In workPart.Annotations

            If Not IsAnnotationInView(ann, dView) Then Continue For

            Dim tagStr As String = ann.Tag.ToString()

            For Each res As CompareResult In results

                If res.DimName.Contains(tagStr) Then

                    Dim dispProps As New DisplayModification
                    dispProps = theSession.DisplayManager.NewDisplayModification()

                    Select Case res.Status
                        Case "PASS"
                            dispProps.ApplyToAllFaces = False
                            dispProps.Color = GREEN
                        Case "FAIL", "MISSING"
                            dispProps.Color = RED
                        Case "EXTRA"
                            dispProps.Color = YELLOW
                    End Select

                    Dim objArr(0) As DisplayableObject
                    objArr(0) = TryCast(ann, DisplayableObject)
                    If objArr(0) IsNot Nothing Then
                        dispProps.Apply(objArr)
                    End If

                    dispProps.Dispose()
                    Exit For

                End If

            Next

        Next

        theSession.Parts.Work.Views.WorkView.UpdateDisplay()

    End Sub

    '=========================================================================
    ' STEP 6 - Generate Comparison Report
    '=========================================================================
    Sub GenerateReport(dView As Drawings.DraftingView,
                       results As List(Of CompareResult))

        Dim passCount    As Integer = 0
        Dim failCount    As Integer = 0
        Dim missingCount As Integer = 0
        Dim extraCount   As Integer = 0

        Dim reportLines As New List(Of String)

        reportLines.Add("=======================================================")
        reportLines.Add("   2D DRAWING vs DESIGN MASTER - COMPARISON REPORT")
        reportLines.Add("=======================================================")
        reportLines.Add("View      : " & dView.Name)
        reportLines.Add("Part      : " & workPart.Name)
        reportLines.Add("Date      : " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        reportLines.Add("-------------------------------------------------------")
        reportLines.Add(String.Format("{0,-25} {1,-12} {2,-12} {3,-10} {4}",
                         "Dimension", "Master Val", "Drawing Val", "Status", "Remarks"))
        reportLines.Add("-------------------------------------------------------")

        For Each res As CompareResult In results

            Select Case res.Status
                Case "PASS"    : passCount    += 1
                Case "FAIL"    : failCount    += 1
                Case "MISSING" : missingCount += 1
                Case "EXTRA"   : extraCount   += 1
            End Select

            reportLines.Add(String.Format("{0,-25} {1,-12} {2,-12} {3,-10} {4}",
                res.DimName.Substring(0, Math.Min(24, res.DimName.Length)),
                res.MasterValue.ToString("F3"),
                res.ActualValue.ToString("F3"),
                res.Status,
                res.Remarks))

        Next

        reportLines.Add("-------------------------------------------------------")
        reportLines.Add("SUMMARY:")
        reportLines.Add("  PASS    : " & passCount)
        reportLines.Add("  FAIL    : " & failCount)
        reportLines.Add("  MISSING : " & missingCount)
        reportLines.Add("  EXTRA   : " & extraCount)
        reportLines.Add("  TOTAL   : " & results.Count)
        reportLines.Add("=======================================================")

        ' Write report to file
        Dim reportPath As String = "C:\NX_Reports\DrawingComparison_" &
                                    dView.Name & "_" &
                                    DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"

        Try
            System.IO.Directory.CreateDirectory("C:\NX_Reports")
            System.IO.File.WriteAllLines(reportPath, reportLines.ToArray())
            MsgBox("Report saved to: " & reportPath & vbNewLine & vbNewLine &
                   "PASS: " & passCount & "  |  FAIL: " & failCount &
                   "  |  MISSING: " & missingCount & "  |  EXTRA: " & extraCount)
        Catch ex As Exception
            ' If file write fails, show in messagebox
            MsgBox(String.Join(vbNewLine, reportLines.ToArray()))
        End Try

    End Sub

    '=========================================================================
    ' HELPER - Check if annotation belongs to a specific view
    '=========================================================================
    Function IsAnnotationInView(ann As Annotation,
                                 dView As Drawings.DraftingView) As Boolean
        Try
            Dim annView As Drawings.DraftingView = 
                TryCast(ann.GetGeometricView(), Drawings.DraftingView)
            If annView IsNot Nothing Then
                Return annView.Tag = dView.Tag
            End If
        Catch
        End Try
        Return False
    End Function

    '=========================================================================
    ' HELPER - Get numeric value from any dimension type
    '=========================================================================
    Function GetDimensionValue(ann As Annotation) As Double
        Try
            Dim dimText As String = ann.GetDimensionText()
            Dim cleaned As String = System.Text.RegularExpressions.Regex.Replace(
                                        dimText, "[^0-9\.\-]", "")
            Dim val As Double = 0
            Double.TryParse(cleaned, val)
            Return val
        Catch
            Return 0
        End Try
    End Function

    '=========================================================================
    ' HELPER - Extract upper and lower tolerance from a dimension
    '=========================================================================
    Sub GetToleranceValues(ann As Annotation,
                           ByRef upperTol As Double,
                           ByRef lowerTol As Double)
        upperTol = 0
        lowerTol = 0
        Try
            Dim dim As Dimension = TryCast(ann, Dimension)
            If dim IsNot Nothing Then
                Dim tolType As DimensionTolerance.ToleranceType = 
                    dim.Tolerance.ToleranceType
                If tolType <> DimensionTolerance.ToleranceType.None Then
                    upperTol = dim.Tolerance.Upper
                    lowerTol = dim.Tolerance.Lower
                End If
            End If
        Catch
        End Try
    End Sub

End Module