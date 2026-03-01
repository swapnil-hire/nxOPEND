Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies

Module BOM_PartList_Balloon

    Dim theSession   As Session   = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI        As UI        = UI.GetUI()
    Dim workPart     As Part      = theSession.Parts.Work

    '=========================================================================
    ' BOM DATA STRUCTURE
    '=========================================================================
    Structure BOMItem
        Dim ItemNumber   As Integer
        Dim PartNumber   As String
        Dim PartName     As String
        Dim Description  As String
        Dim Material     As String
        Dim Quantity     As Integer
        Dim Revision     As String
        Dim CompTag      As Tag        ' NX component tag reference
    End Structure

    '=========================================================================
    ' MAIN
    '=========================================================================
    Sub Main()

        Dim markId As Session.UndoMarkId
        markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "BOM_Creation")

        ' Step 1: Validate environment
        If workPart Is Nothing Then
            MsgBox("No active part found.", MsgBoxStyle.Exclamation, "BOM Creator")
            Return
        End If

        Dim currentSheet As Drawings.DrawingSheet =
                workPart.DrawingSheets.CurrentDrawingSheet
        If currentSheet Is Nothing Then
            MsgBox("No active drawing sheet. Please open a drawing sheet.",
                   MsgBoxStyle.Exclamation, "BOM Creator")
            Return
        End If

        Try
            MsgBox("Starting BOM Creation..." & vbNewLine &
                   "Sheet: " & currentSheet.Name, MsgBoxStyle.Information, "BOM Creator")

            ' Step 2: Collect all components and build BOM list
            Dim bomItems As List(Of BOMItem) = CollectComponents()

            If bomItems.Count = 0 Then
                MsgBox("No components found in the assembly.",
                       MsgBoxStyle.Exclamation, "BOM Creator")
                Return
            End If

            ' Step 3: Create Part List Table on drawing sheet
            Dim partListOrigin As New Point3d(10.0, 10.0, 0.0)  ' Position on sheet (mm)
            CreatePartListTable(currentSheet, bomItems, partListOrigin)

            ' Step 4: Create Balloons on drawing views
            CreateBalloons(currentSheet, bomItems)

            ' Step 5: Show summary
            MsgBox("BOM Created Successfully!" & vbNewLine &
                   "Total Parts : " & bomItems.Count & vbNewLine &
                   "Sheet       : " & currentSheet.Name,
                   MsgBoxStyle.Information, "BOM Creator")

        Catch ex As Exception
            MsgBox("Error: " & ex.Message & vbNewLine & ex.StackTrace,
                   MsgBoxStyle.Critical, "BOM Creator Error")
        End Try

    End Sub

    '=========================================================================
    ' STEP 2 - Collect All Components from Assembly
    '=========================================================================
    Function CollectComponents() As List(Of BOMItem)

        Dim bomList    As New List(Of BOMItem)
        Dim itemNumber As Integer = 1

        ' Dictionary to track quantities (part number -> index in bomList)
        Dim partIndex As New Dictionary(Of String, Integer)

        Try
            ' Traverse all components in the assembly
            Dim rootComp As Component = workPart.ComponentAssembly.RootComponent

            If rootComp Is Nothing Then
                ' Single part (no assembly), add the work part itself
                Dim singleItem As New BOMItem
                singleItem.ItemNumber  = 1
                singleItem.PartNumber  = workPart.Name
                singleItem.PartName    = GetPartAttribute(workPart, "NAME")
                singleItem.Description = GetPartAttribute(workPart, "DESCRIPTION")
                singleItem.Material    = GetPartAttribute(workPart, "MATERIAL")
                singleItem.Revision    = GetPartAttribute(workPart, "DB_PART_REV")
                singleItem.Quantity    = 1
                singleItem.CompTag     = Tag.Null
                bomList.Add(singleItem)
                Return bomList
            End If

            ' Recurse through assembly tree
            TraverseComponents(rootComp, bomList, partIndex, itemNumber)

        Catch ex As Exception
            MsgBox("Error collecting components: " & ex.Message,
                   MsgBoxStyle.Critical, "BOM Error")
        End Try

        Return bomList

    End Function

    '=========================================================================
    ' Recursive traversal of component tree
    '=========================================================================
    Sub TraverseComponents(parentComp   As Component,
                           bomList      As List(Of BOMItem),
                           partIndex    As Dictionary(Of String, Integer),
                           ByRef itemNo As Integer)

        For Each childComp As Component In parentComp.GetChildren()

            Dim compPart As Part = TryCast(childComp.Prototype, Part)
            If compPart Is Nothing Then Continue For

            ' Get part number - try attributes first, fallback to part name
            Dim partNum  As String = GetPartAttribute(compPart, "DB_PART_NO")
            If String.IsNullOrEmpty(partNum) Then partNum = compPart.Name

            ' Check if this part number already exists in BOM (increment qty)
            If partIndex.ContainsKey(partNum) Then
                Dim idx As Integer = partIndex(partNum)
                Dim existing As BOMItem = bomList(idx)
                existing.Quantity += 1
                bomList(idx) = existing
            Else
                ' New part - add to BOM
                Dim newItem As New BOMItem
                newItem.ItemNumber  = itemNo
                newItem.PartNumber  = partNum
                newItem.PartName    = GetPartAttribute(compPart, "NAME")
                newItem.Description = GetPartAttribute(compPart, "DESCRIPTION")
                newItem.Material    = GetPartAttribute(compPart, "MATERIAL")
                newItem.Revision    = GetPartAttribute(compPart, "DB_PART_REV")
                newItem.Quantity    = 1
                newItem.CompTag     = childComp.Tag

                If String.IsNullOrEmpty(newItem.PartName)    Then newItem.PartName    = compPart.Name
                If String.IsNullOrEmpty(newItem.Description) Then newItem.Description = "-"
                If String.IsNullOrEmpty(newItem.Material)    Then newItem.Material    = "-"
                If String.IsNullOrEmpty(newItem.Revision)    Then newItem.Revision    = "A"

                partIndex(partNum) = bomList.Count
                bomList.Add(newItem)
                itemNo += 1
            End If

        Next

    End Sub

    '=========================================================================
    ' STEP 3 - Create Part List Table on Drawing Sheet
    '=========================================================================
    Sub CreatePartListTable(sheet    As Drawings.DrawingSheet,
                            bomItems As List(Of BOMItem),
                            origin   As Point3d)

        ' ---------------------------------------------------------------
        ' Table layout settings
        ' ---------------------------------------------------------------
        Dim colWidths() As Double = {15.0, 35.0, 50.0, 40.0, 30.0, 15.0, 15.0}
        '                            Item  PartNo  Name  Desc  Mat   Qty   Rev
        Dim rowHeight   As Double   = 8.0
        Dim textHeight  As Double   = 3.5
        Dim tableWidth  As Double   = 0
        For Each w As Double In colWidths : tableWidth += w : Next

        Dim headers() As String = {"ITEM", "PART NO.", "PART NAME",
                                    "DESCRIPTION", "MATERIAL", "QTY", "REV"}

        ' ---------------------------------------------------------------
        ' Build table using NX Table Note (TabularNote)
        ' ---------------------------------------------------------------
        Dim tabNoteBuilder As Annotations.TableNoteBuilder =
                workPart.Annotations.CreateTableNoteBuilder(Nothing)

        tabNoteBuilder.AnchorLocation = Annotations.TableNoteBuilder.AnchorType.BottomLeft

        ' Set number of rows = header + BOM rows
        Dim totalRows As Integer = bomItems.Count + 2   ' +1 header, +1 title row

        tabNoteBuilder.RowsAndColumns.SetSize(totalRows, headers.Length)

        ' Set column widths
        For c As Integer = 0 To headers.Length - 1
            tabNoteBuilder.RowsAndColumns.SetColumnWidth(c, colWidths(c))
        Next

        ' Set row heights
        For r As Integer = 0 To totalRows - 1
            tabNoteBuilder.RowsAndColumns.SetRowHeight(r, rowHeight)
        Next

        ' ---------------------------------------------------------------
        ' Row 0 - Title Row (merged)
        ' ---------------------------------------------------------------
        Dim titleCell As Annotations.TableNoteCell =
                tabNoteBuilder.RowsAndColumns.GetCell(0, 0)
        titleCell.Text.SetText(New String() {"BILL OF MATERIALS - " & workPart.Name})
        tabNoteBuilder.RowsAndColumns.MergeCells(0, 0, 0, headers.Length - 1)

        ' ---------------------------------------------------------------
        ' Row 1 - Header Row
        ' ---------------------------------------------------------------
        For c As Integer = 0 To headers.Length - 1
            Dim hCell As Annotations.TableNoteCell =
                    tabNoteBuilder.RowsAndColumns.GetCell(1, c)
            hCell.Text.SetText(New String() {headers(c)})
        Next

        ' ---------------------------------------------------------------
        ' Rows 2+ - BOM Data Rows (filled in REVERSE for BOM convention)
        ' ---------------------------------------------------------------
        For i As Integer = 0 To bomItems.Count - 1

            Dim rowIdx  As Integer = totalRows - 1 - i   ' Bottom-up fill
            Dim item    As BOMItem = bomItems(i)

            Dim cellData() As String = {
                item.ItemNumber.ToString(),
                item.PartNumber,
                item.PartName,
                item.Description,
                item.Material,
                item.Quantity.ToString(),
                item.Revision
            }

            For c As Integer = 0 To cellData.Length - 1
                Dim dCell As Annotations.TableNoteCell =
                        tabNoteBuilder.RowsAndColumns.GetCell(rowIdx, c)
                dCell.Text.SetText(New String() {cellData(c)})
            Next

        Next

        ' ---------------------------------------------------------------
        ' Set origin position on drawing sheet
        ' ---------------------------------------------------------------
        tabNoteBuilder.Origin.Anchor = Annotations.OriginBuilder.LocationType.AbsoluteXy
        tabNoteBuilder.Origin.XValue = origin.X
        tabNoteBuilder.Origin.YValue = origin.Y

        ' Commit the table
        Dim tableNote As Annotations.TableNote =
                CType(tabNoteBuilder.Commit(), Annotations.TableNote)

        tabNoteBuilder.Destroy()

        MsgBox("Part List Table created at (" & origin.X & ", " & origin.Y & ")",
               MsgBoxStyle.Information, "BOM")

    End Sub

    '=========================================================================
    ' STEP 4 - Create Balloons on Drawing Views
    '=========================================================================
    Sub CreateBalloons(sheet    As Drawings.DrawingSheet,
                       bomItems As List(Of BOMItem))

        ' Get all drafting views on the current sheet
        Dim sheetViews As New List(Of Drawings.DraftingView)

        For Each dv As Drawings.DraftingView In workPart.DraftingViews
            If dv.Sheet.Tag = sheet.Tag Then
                sheetViews.Add(dv)
            End If
        Next

        If sheetViews.Count = 0 Then
            MsgBox("No drawing views found on sheet. Balloons skipped.",
                   MsgBoxStyle.Exclamation, "Balloon")
            Return
        End If

        ' Use the first view (typically front/main view) for balloons
        Dim mainView As Drawings.DraftingView = sheetViews(0)

        Dim balloonCount As Integer = 0
        Dim yOffset      As Double  = 0.0

        For Each item As BOMItem In bomItems

            Try
                ' Find component in the view and get its screen position
                Dim balloonPos   As Point3d = GetComponentPositionInView(item.CompTag, mainView, yOffset)
                Dim leaderEndPos As Point3d = balloonPos
                leaderEndPos.X += 20.0   ' Offset balloon symbol from part

                ' -------------------------------------------------------
                ' Create Balloon using ID Symbol builder
                ' -------------------------------------------------------
                Dim idSymBuilder As Annotations.IdSymbolBuilder =
                        workPart.Annotations.CreateIdSymbolBuilder(Nothing)

                ' Set balloon shape - Circle is standard BOM balloon
                idSymBuilder.Type        = Annotations.IdSymbolBuilder.SymbolType.Circle
                idSymBuilder.UpperText   = item.ItemNumber.ToString()
                idSymBuilder.LowerText   = ""
                idSymBuilder.Size        = 8.0   ' balloon diameter in mm

                ' Set balloon position
                idSymBuilder.Origin.Anchor = Annotations.OriginBuilder.LocationType.AbsoluteXy
                idSymBuilder.Origin.XValue = leaderEndPos.X
                idSymBuilder.Origin.YValue = leaderEndPos.Y

                ' Add leader line pointing to component
                Dim leaderBuilder As Annotations.LeaderBuilder =
                        idSymBuilder.CreateLeader()

                leaderBuilder.Attachment = Annotations.LeaderBuilder.LeaderType.OnObject

                Dim leaderPoint As New Point3d(balloonPos.X, balloonPos.Y, 0.0)
                leaderBuilder.SetValue(leaderPoint)

                ' Associate balloon to the component
                If item.CompTag <> Tag.Null Then
                    Dim compObj As NXObject = 
                            CType(theSession.GetObjectManager().GetTaggedObject(item.CompTag), NXObject)
                    idSymBuilder.AssociatedObject = compObj
                End If

                ' Commit the balloon
                Dim balloon As Annotations.IdSymbol =
                        CType(idSymBuilder.Commit(), Annotations.IdSymbol)

                idSymBuilder.Destroy()
                balloonCount += 1
                yOffset += 15.0   ' Space out balloons vertically

            Catch ex As Exception
                ' Skip balloon if component position not found
                Continue For
            End Try

        Next

        MsgBox("Balloons created: " & balloonCount, MsgBoxStyle.Information, "Balloon")

    End Sub

    '=========================================================================
    ' HELPER - Get component visual center position in a drawing view
    '=========================================================================
    Function GetComponentPositionInView(compTag  As Tag,
                                         dView    As Drawings.DraftingView,
                                         yOffset  As Double) As Point3d

        Dim pos As New Point3d(0.0, 0.0, 0.0)

        Try
            If compTag <> Tag.Null Then

                ' Get bounding box of the component in model space
                Dim minCorner(2) As Double
                Dim maxCorner(2) As Double
                Dim compObj As NXObject =
                        CType(theSession.GetObjectManager().GetTaggedObject(compTag), NXObject)

                theUfSession.Modl.AskBoundingBoxExact(compTag, Tag.Null, minCorner, maxCorner)

                ' Compute model space centroid
                Dim cx As Double = (minCorner(0) + maxCorner(0)) / 2.0
                Dim cy As Double = (minCorner(1) + maxCorner(1)) / 2.0
                Dim cz As Double = (minCorner(2) + maxCorner(2)) / 2.0

                ' Map model centroid to drawing view sheet coordinates
                Dim modelPt(2) As Double
                modelPt(0) = cx : modelPt(1) = cy : modelPt(2) = cz

                Dim sheetPt(2) As Double
                theUfSession.Draw.MapModelToDrawing(dView.Tag, modelPt, sheetPt)

                pos.X = sheetPt(0)
                pos.Y = sheetPt(1)
                pos.Z = 0.0

            Else
                ' Fallback: spread balloons at a default location
                pos.X = dView.Origin.X + 30.0
                pos.Y = dView.Origin.Y + 30.0 + yOffset
            End If

        Catch
            ' Fallback position if mapping fails
            pos.X = dView.Origin.X + 30.0
            pos.Y = dView.Origin.Y + 30.0 + yOffset
        End Try

        Return pos

    End Function

    '=========================================================================
    ' HELPER - Read part attribute safely
    '=========================================================================
    Function GetPartAttribute(p As Part, attrName As String) As String
        Try
            Dim attrInfo As NXObject.AttributeInformation =
                    p.GetUserAttribute(attrName, NXObject.AttributeType.String, -1)
            Return attrInfo.StringValue
        Catch
            Try
                ' Try title block / DB attribute
                Dim val As String = Nothing
                theUfSession.Attr.ReadValueAsString(p.Tag, attrName, val)
                Return If(val, "")
            Catch
                Return ""
            End Try
        End Try
    End Function

    '=========================================================================
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

End Module
