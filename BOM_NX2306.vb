Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports NXOpen.Annotations
Imports NXOpen.PDM

Module BOM_TC_NX2306

    Dim theSession   As Session   = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI        As UI        = UI.GetUI()
    Dim workPart     As Part      = theSession.Parts.Work

    '=========================================================================
    ' TEAMCENTER ATTRIBUTE NAMES
    ' These are standard TC integrated NX attribute names
    ' Modify these constants if your TC has custom attribute names
    '=========================================================================
    Const TC_ITEM_ID       As String = "DB_PART_NO"       ' TC Item ID / Part Number
    Const TC_PART_NAME     As String = "COMPONENT_NAME"   ' TC Item Name
    Const TC_OBJECT_NAME   As String = "OBJECT_NAME"      ' TC Object Name fallback
    Const TC_OBJECT_DESC   As String = "OBJECT_DESC"      ' TC Description
    Const TC_REVISION      As String = "LAST_MOD_DATE"    ' TC Revision

    '=========================================================================
    ' BOM DATA STRUCTURE - 4 Columns only
    '=========================================================================
    Structure BOMItem
        Dim ItemNumber As Integer
        Dim PartNumber As String
        Dim PartName   As String
        Dim Quantity   As Integer
        Dim CompTag    As Tag
    End Structure

    '=========================================================================
    ' MAIN
    '=========================================================================
    Sub Main()

        Dim markId As Session.UndoMarkId
        markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "BOM_TC_2306")

        ' --- Validate Part ---
        If workPart Is Nothing Then
            MsgBox("No active part found.",
                   MsgBoxStyle.Exclamation, "BOM TC NX2306")
            Return
        End If

        ' --- Validate Drawing Sheet ---
        Dim currentSheet As DrawingSheet =
                workPart.DrawingSheets.CurrentDrawingSheet

        If currentSheet Is Nothing Then
            MsgBox("No active drawing sheet found." & vbNewLine &
                   "Please open a drawing sheet first.",
                   MsgBoxStyle.Exclamation, "BOM TC NX2306")
            Return
        End If

        Try

            ' Step 1: Collect components from TC integrated assembly
            Dim bomItems As List(Of BOMItem) = CollectTC_Components()

            If bomItems.Count = 0 Then
                MsgBox("No components found." & vbNewLine &
                       "Please check assembly structure in Teamcenter.",
                       MsgBoxStyle.Exclamation, "BOM TC NX2306")
                Return
            End If

            ' Show what was collected for verification
            Dim preview As String = "Components found: " & bomItems.Count & vbNewLine &
                                    "─────────────────────────────────" & vbNewLine
            For Each b As BOMItem In bomItems
                preview &= b.ItemNumber & " | " &
                           b.PartNumber & " | " &
                           b.PartName   & " | Qty:" &
                           b.Quantity   & vbNewLine
            Next

            Dim confirm As MsgBoxResult =
                MsgBox(preview & vbNewLine & "Proceed to create Part List and Balloons?",
                       MsgBoxStyle.YesNo, "BOM Preview - Confirm")

            If confirm = MsgBoxResult.No Then Return

            ' Step 2: Create Part List Table (4 columns)
            CreatePartList_4Col(currentSheet, bomItems)

            ' Step 3: Create Balloons
            CreateBalloons_TC(currentSheet, bomItems)

            MsgBox("BOM Created Successfully!" & vbNewLine &
                   "Items : " & bomItems.Count & vbNewLine &
                   "Sheet : " & currentSheet.Name,
                   MsgBoxStyle.Information, "BOM Complete")

        Catch ex As Exception
            MsgBox("Error: " & ex.Message & vbNewLine & vbNewLine &
                   ex.StackTrace,
                   MsgBoxStyle.Critical, "BOM Error")
        End Try

    End Sub

    '=========================================================================
    ' STEP 1 - Collect Components from Teamcenter Integrated Assembly
    '=========================================================================
    Function CollectTC_Components() As List(Of BOMItem)

        Dim bomList  As New List(Of BOMItem)
        Dim partDict As New Dictionary(Of String, Integer)
        Dim itemNo   As Integer = 1

        Try
            Dim rootComp As Component =
                    workPart.ComponentAssembly.RootComponent

            If rootComp Is Nothing Then
                ' Single part
                Dim single As New BOMItem
                single.ItemNumber = 1
                single.PartNumber = GetTC_PartNumber(workPart)
                single.PartName   = GetTC_PartName(workPart)
                single.Quantity   = 1
                single.CompTag    = Tag.Null
                bomList.Add(single)
                Return bomList
            End If

            ' Traverse only TOP level children (direct children of root)
            ' Change to recursive call if you need full BOM explosion
            For Each child As Component In rootComp.GetChildren()

                Dim cPart As Part = TryCast(child.Prototype, Part)
                If cPart Is Nothing Then Continue For

                ' Get Part Number from Teamcenter
                Dim pNo As String = GetTC_PartNumber(cPart)

                If partDict.ContainsKey(pNo) Then
                    ' Duplicate - increment quantity
                    Dim idx   As Integer = partDict(pNo)
                    Dim exist As BOMItem = bomList(idx)
                    exist.Quantity += 1
                    bomList(idx)    = exist
                Else
                    ' New part
                    Dim newItem As New BOMItem
                    newItem.ItemNumber = itemNo
                    newItem.PartNumber = pNo
                    newItem.PartName   = GetTC_PartName(cPart)
                    newItem.Quantity   = 1
                    newItem.CompTag    = child.Tag

                    partDict(pNo) = bomList.Count
                    bomList.Add(newItem)
                    itemNo += 1
                End If

            Next

        Catch ex As Exception
            MsgBox("Component collection error: " & ex.Message,
                   MsgBoxStyle.Critical, "BOM")
        End Try

        Return bomList

    End Function

    '=========================================================================
    ' STEP 2 - Create 4 Column Part List using NX 2306 PartListBuilder
    '          with Teamcenter attribute mapping
    '=========================================================================
    Sub CreatePartList_4Col(sheet    As DrawingSheet,
                             bomItems As List(Of BOMItem))

        Dim plBuilder As PartListBuilder = Nothing

        Try

            plBuilder = workPart.Annotations.CreatePartListBuilder(Nothing)

            ' --- BOM Level Settings ---
            plBuilder.Settings.Associative = True
            plBuilder.Settings.BomLevel    =
                PartListBuilder.PartListBomLevelOption.TopLevelOnly

            ' --- Sorting: By Item Number ---
            plBuilder.Settings.SortingLevel =
                PartListBuilder.PartListSortingLevelOption.Level1

            ' --- Row height ---
            plBuilder.Settings.RowHeight = 8.0

            ' --- Position on sheet - bottom right (standard for TC drawings) ---
            plBuilder.Origin.Anchor  = OriginBuilder.LocationType.BottomRight
            plBuilder.Origin.XValue  = 400.0   ' Adjust to your title block X
            plBuilder.Origin.YValue  =  10.0   ' Adjust to your title block Y

            ' --- Configure 4 Columns ---
            Configure_4Columns(plBuilder)

            ' --- Commit ---
            Dim partList As PartList = CType(plBuilder.Commit(), PartList)

            If partList IsNot Nothing Then
                MsgBox("Part List created: " &
                       partList.GetAllMemberRows().Length & " rows.",
                       MsgBoxStyle.Information, "Part List")
            End If

        Catch ex As Exception
            MsgBox("Part List creation error: " & ex.Message,
                   MsgBoxStyle.Critical, "Part List")
        Finally
            If plBuilder IsNot Nothing Then plBuilder.Destroy()
        End Try

    End Sub

    '-------------------------------------------------------------------------
    ' 4 Column Configuration mapped to Teamcenter attributes
    '-------------------------------------------------------------------------
    Sub Configure_4Columns(plBuilder As PartListBuilder)

        ' ---------------------------------------------------------------
        ' COLUMN 1 - ITEM NUMBER
        ' NX 2306 TC: FindNumber is auto-generated and linked to balloon
        ' ---------------------------------------------------------------
        Dim col1 As PartListColumnBuilder =
                plBuilder.Columns.CreateColumnBuilder(Nothing)

        col1.ColumnType = PartListColumnBuilder.ColumnTypeOption.FindNumber
        col1.Heading.Text.SetText(New String() {"ITEM"})
        col1.Width       = 15.0
        col1.Alignment   = PartListColumnBuilder.ColumnAlignmentOption.Center
        col1.Commit()

        ' ---------------------------------------------------------------
        ' COLUMN 2 - PART NUMBER (TC: Item ID = DB_PART_NO)
        ' ---------------------------------------------------------------
        Dim col2 As PartListColumnBuilder =
                plBuilder.Columns.CreateColumnBuilder(Nothing)

        col2.ColumnType    = PartListColumnBuilder.ColumnTypeOption.Attribute
        col2.AttributeName = TC_ITEM_ID         ' "DB_PART_NO" from Teamcenter
        col2.Heading.Text.SetText(New String() {"PART NUMBER"})
        col2.Width         = 50.0
        col2.Alignment     = PartListColumnBuilder.ColumnAlignmentOption.Left
        col2.Commit()

        ' ---------------------------------------------------------------
        ' COLUMN 3 - QUANTITY
        ' NX 2306 TC: Quantity is automatically counted from instances
        ' ---------------------------------------------------------------
        Dim col3 As PartListColumnBuilder =
                plBuilder.Columns.CreateColumnBuilder(Nothing)

        col3.ColumnType = PartListColumnBuilder.ColumnTypeOption.Quantity
        col3.Heading.Text.SetText(New String() {"QTY"})
        col3.Width       = 15.0
        col3.Alignment   = PartListColumnBuilder.ColumnAlignmentOption.Center
        col3.Commit()

        ' ---------------------------------------------------------------
        ' COLUMN 4 - PART NAME (TC: OBJECT_NAME or COMPONENT_NAME)
        ' ---------------------------------------------------------------
        Dim col4 As PartListColumnBuilder =
                plBuilder.Columns.CreateColumnBuilder(Nothing)

        col4.ColumnType    = PartListColumnBuilder.ColumnTypeOption.Attribute
        col4.AttributeName = TC_PART_NAME       ' "COMPONENT_NAME" from Teamcenter
        col4.Heading.Text.SetText(New String() {"PART NAME"})
        col4.Width         = 70.0
        col4.Alignment     = PartListColumnBuilder.ColumnAlignmentOption.Left
        col4.Commit()

    End Sub

    '=========================================================================
    ' STEP 3 - Create Balloons linked to Part List (Teamcenter NX 2306)
    '=========================================================================
    Sub CreateBalloons_TC(sheet    As DrawingSheet,
                           bomItems As List(Of BOMItem))

        ' Get views on this sheet
        Dim sheetViews As New List(Of DraftingView)
        For Each dv As DraftingView In workPart.DraftingViews
            If dv.Sheet IsNot Nothing AndAlso
               dv.Sheet.Tag = sheet.Tag Then
                sheetViews.Add(dv)
            End If
        Next

        If sheetViews.Count = 0 Then
            MsgBox("No views on sheet. Balloons skipped.",
                   MsgBoxStyle.Exclamation, "Balloon")
            Return
        End If

        ' --- Try AutoBalloon first ---
        If TryAutoBalloon_TC() Then Return

        ' --- Fallback: Manual balloon per component ---
        ManualBalloons_TC(sheetViews(0), bomItems)

    End Sub

    '-------------------------------------------------------------------------
    ' AutoBalloon for TC NX 2306
    '-------------------------------------------------------------------------
    Function TryAutoBalloon_TC() As Boolean

        Try

            Dim abBuilder As AutoBalloonBuilder =
                    workPart.Annotations.CreateAutoBalloonBuilder()

            ' Standard circular balloon
            abBuilder.BalloonType =
                AutoBalloonBuilder.BalloonShapeOption.CircularBalloon

            abBuilder.Size = 8.0

            ' TC NX 2306: Attach to silhouette edge of component
            abBuilder.AttachmentType =
                AutoBalloonBuilder.AttachmentTypeOption.SilhouetteEdge

            ' Don't re-balloon already ballooned items
            abBuilder.IgnoreExistingBalloons = True

            ' TC NX 2306: Link balloon number to Part List find number
            abBuilder.BalloonType =
                AutoBalloonBuilder.BalloonShapeOption.CircularBalloon

            abBuilder.Commit()
            abBuilder.Destroy()

            MsgBox("Auto Balloons created and linked to Part List.",
                   MsgBoxStyle.Information, "Balloon")
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    '-------------------------------------------------------------------------
    ' Manual Balloon placement for TC NX 2306
    '-------------------------------------------------------------------------
    Sub ManualBalloons_TC(mainView As DraftingView,
                           bomItems As List(Of BOMItem))

        Dim count   As Integer = 0
        Dim yOffset As Double  = 0.0

        For Each item As BOMItem In bomItems

            Try

                ' Get component center mapped to sheet coords
                Dim compPt As Point3d =
                        MapComponentToSheet(item.CompTag, mainView, yOffset)

                ' Balloon position offset from component
                Dim symPt As New Point3d(compPt.X + 25.0,
                                          compPt.Y + 10.0,
                                          0.0)

                ' --- Create Balloon ---
                Dim idBuilder As IdSymbolBuilder =
                        workPart.Annotations.CreateIdSymbolBuilder(Nothing)

                idBuilder.Type      = IdSymbolBuilder.SymbolType.Circle
                idBuilder.UpperText = item.ItemNumber.ToString()
                idBuilder.LowerText = ""
                idBuilder.Size      = 8.0

                ' Position
                idBuilder.Origin.Anchor  = OriginBuilder.LocationType.AbsoluteXy
                idBuilder.Origin.XValue  = symPt.X
                idBuilder.Origin.YValue  = symPt.Y

                ' Leader line from balloon to component
                Dim leaderPts(0) As Point3d
                leaderPts(0) = New Point3d(compPt.X, compPt.Y, 0.0)
                idBuilder.Leader.Leaders.Item(0).SetLeaderPoints(leaderPts)

                ' TC NX 2306: Associate balloon to component NX object
                If item.CompTag <> Tag.Null Then
                    Dim tagObj As TaggedObject =
                            theSession.GetObjectManager().GetTaggedObject(item.CompTag)
                    Dim nxObj As NXObject = TryCast(tagObj, NXObject)
                    If nxObj IsNot Nothing Then
                        idBuilder.AssociatedObject = nxObj
                    End If
                End If

                idBuilder.Commit()
                idBuilder.Destroy()

                count   += 1
                yOffset += 18.0

            Catch
                Continue For
            End Try

        Next

        MsgBox("Balloons placed: " & count,
               MsgBoxStyle.Information, "Balloon")

    End Sub

    '=========================================================================
    ' HELPER - Get TC Part Number
    ' Priority: DB_PART_NO > OBJECT_NAME > Part file name
    '=========================================================================
    Function GetTC_PartNumber(p As Part) As String

        ' Priority 1: TC standard item ID attribute
        Dim val As String = SafeGetTC_Attr(p, "DB_PART_NO")
        If Not String.IsNullOrEmpty(val) Then Return val

        ' Priority 2: TC object name
        val = SafeGetTC_Attr(p, "OBJECT_NAME")
        If Not String.IsNullOrEmpty(val) Then Return val

        ' Priority 3: UF part name
        Dim partName As String = ""
        Try
            theUfSession.Part.AskPartName(p.Tag, partName)
            If Not String.IsNullOrEmpty(partName) Then Return partName
        Catch
        End Try

        ' Fallback: NX part file name
        Return p.Name

    End Function

    '=========================================================================
    ' HELPER - Get TC Part Name
    ' Priority: COMPONENT_NAME > OBJECT_DESC > OBJECT_NAME > part name
    '=========================================================================
    Function GetTC_PartName(p As Part) As String

        Dim val As String = SafeGetTC_Attr(p, "COMPONENT_NAME")
        If Not String.IsNullOrEmpty(val) Then Return val

        val = SafeGetTC_Attr(p, "OBJECT_DESC")
        If Not String.IsNullOrEmpty(val) Then Return val

        val = SafeGetTC_Attr(p, "OBJECT_NAME")
        If Not String.IsNullOrEmpty(val) Then Return val

        Return p.Name

    End Function

    '=========================================================================
    ' HELPER - Safe Teamcenter Attribute Reader (NX 2306 TC)
    ' Tries all 3 methods NX 2306 TC uses to store attributes
    '=========================================================================
    Function SafeGetTC_Attr(p As Part, attrName As String) As String

        ' Method 1: GetUserAttribute (TC synced string attributes)
        Try
            Dim attr As NXObject.AttributeInformation =
                    p.GetUserAttribute(attrName,
                                       NXObject.AttributeType.String, -1)
            If Not String.IsNullOrEmpty(attr.StringValue) Then
                Return attr.StringValue.Trim()
            End If
        Catch
        End Try

        ' Method 2: UF Attr read (legacy TC attribute method)
        Try
            Dim val As String = ""
            theUfSession.Attr.ReadValueAsString(p.Tag, attrName, val)
            If Not String.IsNullOrEmpty(val) Then Return val.Trim()
        Catch
        End Try

        ' Method 3: Scan all string attributes (handles TC custom attr names)
        Try
            Dim allAttrs() As NXObject.AttributeInformation =
                    p.GetAllAttributesByType(NXObject.AttributeType.String)
            For Each a As NXObject.AttributeInformation In allAttrs
                If a.Title.ToUpper().Trim() = attrName.ToUpper().Trim() Then
                    Return a.StringValue.Trim()
                End If
            Next
        Catch
        End Try

        Return ""

    End Function

    '=========================================================================
    ' HELPER - Diagnostic: List ALL TC attributes on a part
    '          Run this first to discover your exact TC attribute names
    '=========================================================================
    Sub DiagnoseTC_Attributes()

        If workPart Is Nothing Then Return

        Dim msg As String = "TC Attributes on: " & workPart.Name & vbNewLine &
                            "─────────────────────────────" & vbNewLine

        Try
            Dim allAttrs() As NXObject.AttributeInformation =
                    workPart.GetAllAttributesByType(NXObject.AttributeType.String)
            For Each a As NXObject.AttributeInformation In allAttrs
                msg &= a.Title & " = " & a.StringValue & vbNewLine
            Next
        Catch ex As Exception
            msg &= "Error: " & ex.Message
        End Try

        MsgBox(msg, MsgBoxStyle.Information, "TC Attribute Diagnostic")

    End Sub

    '=========================================================================
    ' HELPER - Map 3D component center to 2D sheet coordinates
    '=========================================================================
    Function MapComponentToSheet(compTag As Tag,
                                  dView   As DraftingView,
                                  yOff    As Double) As Point3d

        Dim result As New Point3d(0, 0, 0)

        Try
            If compTag <> Tag.Null Then

                Dim mn(2) As Double
                Dim mx(2) As Double
                theUfSession.Modl.AskBoundingBoxExact(compTag, Tag.Null, mn, mx)

                Dim center(2) As Double
                center(0) = (mn(0) + mx(0)) / 2.0
                center(1) = (mn(1) + mx(1)) / 2.0
                center(2) = (mn(2) + mx(2)) / 2.0

                Dim sheetPt(2) As Double
                theUfSession.Draw.MapModelToDrawing(dView.Tag, center, sheetPt)

                result.X = sheetPt(0)
                result.Y = sheetPt(1)
            Else
                result.X = dView.Origin.X + 20.0
                result.Y = dView.Origin.Y + 20.0 + yOff
            End If
        Catch
            result.X = dView.Origin.X + 20.0
            result.Y = dView.Origin.Y + 20.0 + yOff
        End Try

        Return result

    End Function

    '=========================================================================
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

End Module
