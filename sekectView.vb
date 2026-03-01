Imports System
Imports NXOpen
Imports NXOpen.UF

Module SelectViewExample

    Dim theSession As Session = Session.GetSession()
    Dim theUFSession As UFSession = UFSession.GetUFSession()
    Dim theUI As UI = UI.GetUI()

    Sub Main()

        Dim workPart As Part = theSession.Parts.Work

        If workPart Is Nothing Then
            MsgBox("No active part found. Please open a part or drawing.")
            Return
        End If

        ' Check if we are in a drawing sheet context
        Dim drawingSheet As Drawings.DrawingSheet = workPart.DrawingSheets.CurrentDrawingSheet

        If drawingSheet Is Nothing Then
            MsgBox("No active drawing sheet found. Please open a drawing sheet.")
            Return
        End If

        MsgBox("Active Drawing Sheet: " & drawingSheet.Name)

        ' ---- METHOD 1: Select view interactively using Selection Manager ----
        SelectViewInteractive(workPart)

        ' ---- METHOD 2: Select view by name programmatically ----
        ' SelectViewByName(workPart, "TOP@1")  ' Uncomment and pass your view name

    End Sub

    '-------------------------------------------------------------------------
    ' METHOD 1 - Interactive Selection: User clicks on a view
    '-------------------------------------------------------------------------
    Sub SelectViewInteractive(workPart As Part)

        Dim selectionManager As SelectionManager = theUI.SelectionManager

        ' Define selection options
        Dim selOptions As New Selection.Options
        selOptions.SubjectScope = Selection.Scope.AnyInAssembly

        ' Set mask to select Drawing Views only
        Dim maskTriples(0) As Selection.MaskTriple
        maskTriples(0).Type = UFConstants.UF_drawing_view_type
        maskTriples(0).Subtype = 0
        maskTriples(0).SolidBodySubtype = 0

        Dim selectedObjects() As TaggedObject = Nothing
        Dim cursor As Point3d

        Dim response As Selection.Response = selectionManager.SelectTaggedObjects(
            "Select a Drawing View",       ' Prompt message
            "Select View",                 ' Dialog title
            Selection.Scope.AnyInAssembly,
            Selection.Action.SelectOne,
            False,
            False,
            maskTriples,
            selectedObjects)

        If response = Selection.Response.Ok OrElse response = Selection.Response.ObjectSelected Then
            If selectedObjects IsNot Nothing AndAlso selectedObjects.Length > 0 Then
                Dim selectedView As Drawings.DraftingView = 
                    TryCast(selectedObjects(0), Drawings.DraftingView)

                If selectedView IsNot Nothing Then
                    MsgBox("Selected View: " & selectedView.Name & 
                           vbNewLine & "View Style: " & selectedView.Style.ToString())

                    ' Highlight the selected view
                    selectedView.Highlight()

                    ' You can now work with the view
                    ProcessSelectedView(selectedView)
                Else
                    MsgBox("Selected object is not a Drawing View.")
                End If
            End If
        Else
            MsgBox("No view selected or selection cancelled.")
        End If

    End Sub

    '-------------------------------------------------------------------------
    ' METHOD 2 - Select / Get view by name programmatically
    '-------------------------------------------------------------------------
    Sub SelectViewByName(workPart As Part, viewName As String)

        Dim foundView As Drawings.DraftingView = Nothing

        ' Loop through all drawing views in the part
        For Each dView As Drawings.DraftingView In workPart.DraftingViews

            If dView.Name.ToUpper() = viewName.ToUpper() Then
                foundView = dView
                Exit For
            End If

        Next

        If foundView IsNot Nothing Then
            MsgBox("View Found: " & foundView.Name)

            ' Make the view active/current
            foundView.Activate()

            ' Highlight the view
            foundView.Highlight()

            ' Process it
            ProcessSelectedView(foundView)
        Else
            MsgBox("View with name '" & viewName & "' not found.")
        End If

    End Sub

    '-------------------------------------------------------------------------
    ' METHOD 3 - List ALL views in the current drawing sheet
    '-------------------------------------------------------------------------
    Sub ListAllDrawingViews(workPart As Part)

        Dim viewList As String = "Drawing Views in Current Part:" & vbNewLine & vbNewLine

        Dim count As Integer = 0
        For Each dView As Drawings.DraftingView In workPart.DraftingViews
            count += 1
            viewList &= count.ToString() & ". " & dView.Name & 
                        "  [" & dView.Style.ToString() & "]" & vbNewLine
        Next

        If count = 0 Then
            MsgBox("No drawing views found in this part.")
        Else
            MsgBox(viewList)
        End If

    End Sub

    '-------------------------------------------------------------------------
    ' Process the selected view - extend this for your use case
    '-------------------------------------------------------------------------
    Sub ProcessSelectedView(dView As Drawings.DraftingView)

        ' Get view origin (center point on drawing sheet)
        Dim origin As Point3d = dView.Origin

        ' Get view scale
        Dim scale As Double = dView.Scale

        ' Get view name
        Dim viewName As String = dView.Name

        ' Get drawing sheet the view belongs to
        Dim sheet As Drawings.DrawingSheet = dView.OwningSheet

        Dim info As String = 
            "View Details:" & vbNewLine &
            "  Name     : " & viewName & vbNewLine &
            "  Sheet    : " & sheet.Name & vbNewLine &
            "  Scale    : 1:" & scale.ToString("F3") & vbNewLine &
            "  Origin X : " & origin.X.ToString("F3") & vbNewLine &
            "  Origin Y : " & origin.Y.ToString("F3")

        MsgBox(info)

    End Sub

End Module