' NXOpen VB.NET Journal
' Update Section View Hatching based on Material

Imports System
Imports NXOpen
Imports NXOpen.Drawings
Imports NXOpen.Annotations

Module SectionHatchUpdater

    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim lw As ListingWindow = theSession.ListingWindow

    Sub Main()

        lw.Open()
        lw.WriteLine("=== Section Hatch Updater Started ===")

        Try
            Dim workPart As Part = theSession.Parts.Work

            If workPart Is Nothing Then
                lw.WriteLine("No active part found.")
                Return
            End If

            ' Process all drawing sheets
            For Each sheet As DrawingSheet In workPart.DrawingSheets
                lw.WriteLine("Processing Sheet: " & sheet.Name)
                ProcessSheet(sheet)
            Next

        Catch ex As Exception
            lw.WriteLine("Error: " & ex.Message)
        End Try

        lw.WriteLine("=== Completed ===")
        lw.Close()

    End Sub

    '---------------------------------------------
    ' Process each sheet
    '---------------------------------------------
    Sub ProcessSheet(sheet As DrawingSheet)

        Dim views() As DraftingView = sheet.GetDraftingViews()

        For Each view As DraftingView In views

            If TypeOf view Is SectionView Then
                lw.WriteLine("  Section View Found: " & view.Name)
                ProcessSectionView(CType(view, SectionView))
            End If

        Next

    End Sub

    '---------------------------------------------
    ' Process Section View
    '---------------------------------------------
    Sub ProcessSectionView(secView As SectionView)

        Try
            ' Get visible objects in section view
            Dim objects() As DisplayableObject = secView.AskVisibleObjects()

            For Each obj As DisplayableObject In objects

                Dim body As Body = TryCast(obj, Body)

                If body IsNot Nothing Then

                    Dim matName As String = GetMaterial(body)
                    Dim hatchPattern As String = GetHatchPattern(matName)

                    lw.WriteLine("    Body: " & body.JournalIdentifier & _
                                 " | Material: " & matName & _
                                 " | Hatch: " & hatchPattern)

                    ApplyHatch(secView, body, hatchPattern)

                End If

            Next

        Catch ex As Exception
            lw.WriteLine("Error in SectionView: " & ex.Message)
        End Try

    End Sub

    '---------------------------------------------
    ' Get Material from Body / Attributes
    '---------------------------------------------
    Function GetMaterial(body As Body) As String

        Try
            ' Try NX material first
            If body.Material IsNot Nothing Then
                Return body.Material.Name
            End If

            ' Try attributes
            If body.HasUserAttribute("Material", NXObject.AttributeType.String, -1) Then
                Return body.GetUserAttributeAsString("Material", NXObject.AttributeType.String, -1)
            End If

        Catch ex As Exception
        End Try

        Return "DEFAULT"

    End Function

    '---------------------------------------------
    ' Material → Hatch Mapping
    '---------------------------------------------
    Function GetHatchPattern(materialName As String) As String

        materialName = materialName.ToUpper()

        If materialName.Contains("STEEL") Then
            Return "ANSI31"
        ElseIf materialName.Contains("ALUMINUM") Then
            Return "ANSI32"
        ElseIf materialName.Contains("CAST") Then
            Return "ANSI33"
        ElseIf materialName.Contains("PLASTIC") Then
            Return "ANSI37"
        Else
            Return "ANSI31" ' Default
        End If

    End Function

    '---------------------------------------------
    ' Apply Hatch (Skeleton – extend as needed)
    '---------------------------------------------
    Sub ApplyHatch(secView As SectionView, body As Body, hatchPattern As String)

        Try
            ' NOTE:
            ' NXOpen does not directly allow simple "set hatch per body"
            ' You typically need to work with SectionViewBuilder or HatchBuilder

            Dim workPart As Part = theSession.Parts.Work

            ' Placeholder logic (to be extended)
            ' ----------------------------------
            ' 1. Identify section region
            ' 2. Create or edit hatch region
            ' 3. Assign pattern (ANSI31 etc.)

            lw.WriteLine("      Applying Hatch: " & hatchPattern)

            ' TODO:
            ' Use HatchBuilder or Drafting API
            ' Example (pseudo-real structure):

            'Dim hatchBuilder As Annotations.HatchBuilder = _
            '    workPart.Annotations.Hatches.CreateHatchBuilder(Nothing)

            'hatchBuilder.Pattern = hatchPattern
            'hatchBuilder.Commit()
            'hatchBuilder.Destroy()

        Catch ex As Exception
            lw.WriteLine("Error applying hatch: " & ex.Message)
        End Try

    End Sub

    '---------------------------------------------
    ' Required NX Entry Points
    '---------------------------------------------
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return Session.LibraryUnloadOption.Immediately
    End Function

End Module
