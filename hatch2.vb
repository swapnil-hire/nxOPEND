Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Drawings
Imports NXOpen.Assemblies

Module MaterialHatchAutomator

    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim lw As ListingWindow = theSession.ListingWindow

    ' --- CONFIGURATION: Material to Pattern Mapping ---
    ' Standard NX/ANSI pattern names: ANSI31, ANSI32, etc.
    ReadOnly HatchMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Steel", "ANSI31"},
        {"Aluminum", "ANSI32"},
        {"Iron", "ANSI33"},
        {"Plastic", "ANSI37"},
        {"Rubber", "ANSI38"},
        {"Copper", "ANSI33"},
        {"Default", "ANSI31"}
    }

    Sub Main(ByVal args() As String)
        lw.Open()
        Dim workPart As Part = theSession.Parts.Work

        If workPart Is Nothing Then Return

        ' Undo Mark for safety
        Dim markId As Session.UndoMarkId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Update Section Hatches")

        Try
            ProcessDraftingHatches(workPart)
        Catch ex As Exception
            lw.WriteLine("Error during execution: " & ex.Message)
        End Try
    End Sub

    Sub ProcessDraftingHatches(ByVal workPart As Part)
        Dim currentSheet As DrawingSheet = workPart.DrawingSheets.CurrentDrawingSheet
        If currentSheet Is Nothing Then
            lw.WriteLine("No active drawing sheet found.")
            Return
        End If

        ' Identify all section views on the current sheet
        Dim sectionViews As New List(Of SectionView)
        For Each view As DraftingView In currentSheet.GetDraftingViews()
            If TypeOf view Is SectionView Then
                sectionViews.Add(DirectCast(view, SectionView))
            End If
        Next

        If sectionViews.Count = 0 Then
            lw.WriteLine("No section views found on active sheet.")
            Return
        End If

        ' Iterate through hatches in the part
        Dim hatchCount As Integer = 0
        For Each hatch As Annotations.Hatch In workPart.Annotations.Hatches
            
            ' Only process hatches belonging to our target section views
            Dim hatchView As DraftingView = TryCast(hatch.View, DraftingView)
            If hatchView IsNot Nothing AndAlso sectionViews.Contains(hatchView) Then
                
                ' 1. Determine the Material of the sectioned body
                Dim materialName As String = GetMaterialFromHatch(hatch)
                
                ' 2. Determine correct pattern
                Dim patternName As String = GetPatternForMaterial(materialName)
                
                ' 3. Apply the hatch update
                If ApplyHatchUpdate(workPart, hatch, patternName) Then
                    hatchCount += 1
                End If
            End If
        Next

        lw.WriteLine(String.Format("Successfully updated {0} hatches based on material assignments.", hatchCount))
    End Sub

    ''' <summary>
    ''' Resolves the material attribute from the component/body associated with the hatch.
    ''' </summary>
    Function GetMaterialFromHatch(ByVal hatch As Annotations.Hatch) As String
        Try
            ' Use UF to find the geometry the hatch is associated with
            Dim n_geoms As Integer = 0
            Dim geoms() As Tag = Nothing
            theUfSession.Draw.AskHatchGeom(hatch.Tag, n_geoms, geoms)

            If n_geoms > 0 Then
                ' Get the prototype of the geometry to find attributes
                Dim obj As NXObject = DirectCast(NXObjectManager.Get(geoms(0)), NXObject)
                
                ' If it's a component or body, check attributes
                ' Prioritizing "Material" or "DB_Part_Material" (Common for Teamcenter)
                If obj.HasUserAttribute("Material", NXObject.AttributeType.String, -1) Then
                    Return obj.GetStringAttribute("Material")
                ElseIf obj.HasUserAttribute("DB_Part_Material", NXObject.AttributeType.String, -1) Then
                    Return obj.GetStringAttribute("DB_Part_Material")
                End If
            End If
        Catch ex As Exception
            ' Fallback for single-part files or unlinked geometry
        End Try

        Return "Default"
    End Function

    Function GetPatternForMaterial(ByVal mat As String) As String
        For Each key In HatchMap.Keys
            If mat.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return HatchMap(key)
            End If
        Next
        Return HatchMap("Default")
    End Function

    ''' <summary>
    ''' Re-applies the hatch pattern using the EditSectionHatchBuilder.
    ''' </summary>
    Function ApplyHatchUpdate(ByVal workPart As Part, ByVal hatch As Annotations.Hatch, ByVal pattern As String) As Boolean
        Dim editHatchBuilder As EditSectionHatchBuilder = Nothing
        Try
            editHatchBuilder = workPart.DraftingViews.CreateEditSectionHatchBuilder()
            editHatchBuilder.Hatches.Add(hatch)
            
            ' Update pattern name
            editHatchBuilder.CrossHatchName = pattern
            
            editHatchBuilder.Commit()
            Return True
        Catch ex As Exception
            Return False
        Finally
            If editHatchBuilder IsNot Nothing Then editHatchBuilder.Destroy()
        End Try
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return Session.LibraryUnloadOption.Immediately
    End Function

End Module
