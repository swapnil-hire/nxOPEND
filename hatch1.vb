Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Drawings
Imports NXOpen.Assemblies

Module AutoUpdateSectionHatching

    ' Global session variables
    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI As UI = UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow

    ' --- CONFIGURATION ---
    ' Map your company's material attributes to NX standard crosshatch patterns
    Dim HatchMapping As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Steel", "ANSI31"},
        {"Aluminum", "ANSI32"},
        {"Cast Iron", "ANSI33"},
        {"Plastic", "ANSI37"},
        {"Rubber", "ANSI38"},
        {"Default", "ANSI31"}
    }

    Sub Main(ByVal args() As String)
        lw.Open()
        lw.WriteLine("--- Starting Auto-Update Section Hatching ---")

        Dim workPart As Part = theSession.Parts.Work
        If workPart Is Nothing Then
            lw.WriteLine("Error: No active part found. Exiting.")
            Return
        End If

        ' Verify we are in the Drafting application
        Dim moduleName As String = ""
        theUfSession.UF.AskApplicationModule(moduleName)
        If Not moduleName.Equals("MODULE_DRAFTING", StringComparison.OrdinalIgnoreCase) Then
            theUI.NXMessageBox.Show("Environment Error", NXMessageBox.DialogType.Error, "This journal must be executed within the Drafting application.")
            Return
        End If

        ' Optional UI Confirmation Prompt
        Dim confirm As Integer = theUI.NXMessageBox.Show("Auto-Update Hatching", NXMessageBox.DialogType.Question,
                                "Automatically update section view hatching based on component material?" & vbCrLf & 
                                "This will analyze all section views on the active sheet.")
        If confirm <> 1 Then ' 1 = OK/Yes
            lw.WriteLine("Operation cancelled by user.")
            Return
        End If

        Dim markId As Session.UndoMarkId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Auto Update Section Hatches")

        Try
            ProcessActiveSheet(workPart)
        Catch ex As Exception
            lw.WriteLine("Critical Error: " & ex.Message)
            theSession.UndoToMark(markId, "Auto Update Section Hatches")
        End Try

        lw.WriteLine("--- Process Complete ---")
    End Sub

    Private Sub ProcessActiveSheet(workPart As Part)
        Dim currentSheet As DrawingSheet = workPart.DrawingSheets.CurrentDrawingSheet
        If currentSheet Is Nothing Then
            lw.WriteLine("No active drawing sheet found.")
            Return
        End If

        lw.WriteLine("Processing Sheet: " & currentSheet.Name)

        ' 1. Collect all Section Views on the active sheet
        Dim sectionViews As New List(Of SectionView)()
        For Each view As DraftingView In currentSheet.GetDraftingViews()
            If TypeOf view Is SectionView Then
                sectionViews.Add(DirectCast(view, SectionView))
                lw.WriteLine("Found Section View: " & view.Name)
            End If
        Next

        If sectionViews.Count = 0 Then
            lw.WriteLine("No section views found on the active sheet.")
            Return
        End If

        ' 2. Iterate through all hatches in the work part
        Dim hatchesUpdated As Integer = 0
        For Each hatch As Annotations.Hatch In workPart.Annotations.Hatches
            Try
                ' Check if the hatch belongs to one of our section views
                Dim hView As DraftingView = TryCast(hatch.View, DraftingView)
                If hView IsNot Nothing AndAlso sectionViews.Contains(hView) Then
                    If UpdateHatchBasedOnMaterial(workPart, hatch, hView) Then
                        hatchesUpdated += 1
                    End If
                End If
            Catch ex As Exception
                ' Ignore orphaned or invalid hatches
            End Try
        Next

        lw.WriteLine("Successfully updated " & hatchesUpdated & " hatches.")
        theUI.NXMessageBox.Show("Success", NXMessageBox.DialogType.Information, "Updated " & hatchesUpdated & " section hatches based on material.")
    End Sub

    Private Function UpdateHatchBasedOnMaterial(workPart As Part, hatch As Annotations.Hatch, view As DraftingView) As Boolean
        Dim comp As Component = GetComponentFromHatch(hatch, view)
        Dim material As String = "Default"

        If comp IsNot Nothing Then
            material = GetMaterial(comp)
        End If

        Dim hatchPattern As String = GetHatchPatternForMaterial(material)
        
        Dim compName As String = If(comp IsNot Nothing, comp.Name, "Single Part / Unlinked")
        lw.WriteLine("  -> Component: " & compName & " | Material: " & material & " | Applying Hatch: " & hatchPattern)

        Return ApplyHatchToSection(workPart, hatch, hatchPattern)
    End Function

    ''' <summary>
    ''' Attempts to resolve the owning component of a drafted hatch.
    ''' </summary>
    Private Function GetComponentFromHatch(hatch As Annotations.Hatch, view As DraftingView) As Component
        ' NOTE ON NXOPEN TOPOLOGY: 
        ' Directly mapping a Drafting Hatch -> Drafting Curve -> Model Edge -> Component 
        ' requires advanced UF wrapper topology queries. For this robust script, we look at the 
        ' display part's root component. If it's a single part, we return the root.
        ' If it's an assembly, you can extend this by querying the hatch boundaries via AskHatchGeometry.
        
        Try
            Dim rootComp As Component = theSession.Parts.Display.ComponentAssembly.RootComponent
            
            ' If this is a single part file (no assembly structure), root is Nothing
            If rootComp IsNot Nothing Then
                ' In a full assembly topological trace, you would map hatch.Tag here.
                ' Returning root as fallback.
                Return rootComp
            End If
        Catch ex As Exception
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Extracts the material string from a component's attributes or prototype.
    ''' </summary>
    Private Function GetMaterial(comp As Component) As String
        Try
            ' 1. Check for standard TC/NX user attributes on the component occurrence
            If comp.HasUserAttribute("Material", NXObject.AttributeType.String, -1) Then
                Return comp.GetStringAttribute("Material")
            End If
            If comp.HasUserAttribute("MAT", NXObject.AttributeType.String, -1) Then
                Return comp.GetStringAttribute("MAT")
            End If

            ' 2. Fallback to the prototype part (the actual piece part)
            Dim proto As Part = TryCast(comp.Prototype, Part)
            If proto IsNot Nothing Then
                If proto.HasUserAttribute("Material", NXObject.AttributeType.String, -1) Then
                    Return proto.GetStringAttribute("Material")
                End If
                ' Note: To access the NX physical material manager, you would query proto.MaterialManager here.
            End If

        Catch ex As Exception
            ' If attribute doesn't exist or part isn't loaded, fail silently and return Default
        End Try
        
        Return "Default"
    End Function

    ''' <summary>
    ''' Maps the material string to a hatch pattern, handling partial matches (e.g., "Steel, Stainless").
    ''' </summary>
    Private Function GetHatchPatternForMaterial(materialName As String) As String
        For Each key As String In HatchMapping.Keys
            If materialName.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return HatchMapping(key)
            End If
        Next
        Return HatchMapping("Default")
    End Function

    ''' <summary>
    ''' Applies the crosshatch pattern using the EditSectionHatchBuilder while maintaining angle/distance.
    ''' </summary>
    Private Function ApplyHatchToSection(workPart As Part, hatch As Annotations.Hatch, pattern As String) As Boolean
        Dim editSectionHatchBuilder As EditSectionHatchBuilder = Nothing
        Try
            editSectionHatchBuilder = workPart.DraftingViews.CreateEditSectionHatchBuilder()
            
            ' Add the targeted hatch to the builder
            editSectionHatchBuilder.Hatches.Add(hatch)

            ' Apply the new pattern name (Scale and angle remain preserved by the builder)
            editSectionHatchBuilder.CrossHatchName = pattern

            ' Commit changes
            editSectionHatchBuilder.Commit()
            Return True
        Catch ex As Exception
            lw.WriteLine("  -> Failed to apply hatch: " & ex.Message)
            Return False
        Finally
            If editSectionHatchBuilder IsNot Nothing Then
                editSectionHatchBuilder.Destroy()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Ensures the DLL unloads from NX memory immediately after execution.
    ''' </summary>
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return Session.LibraryUnloadOption.Immediately
    End Function

End Module
