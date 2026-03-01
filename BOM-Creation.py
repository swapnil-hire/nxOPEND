# =============================================================================
# BOM Creation - Part List & Balloon
# NX 2306 | Teamcenter Integrated | Python NXOpen
# 4 Columns: Item Number | Part Number | QTY | Part Name
# =============================================================================

import NXOpen
import NXOpen.UF
import NXOpen.Assemblies
import NXOpen.Annotations
import NXOpen.PDM
from dataclasses import dataclass, field
from typing import List, Dict, Optional

# =============================================================================
# SESSION INIT
# =============================================================================
the_session    = NXOpen.Session.GetSession()
the_uf_session = NXOpen.UF.UFSession.GetUFSession()
the_ui         = NXOpen.UI.GetUI()
work_part      = the_session.Parts.Work

# =============================================================================
# TEAMCENTER ATTRIBUTE CONSTANTS
# Modify these to match your TC attribute names
# Run diagnose_tc_attributes() first to discover exact names
# =============================================================================
TC_ITEM_ID    = "DB_PART_NO"       # TC Item ID / Part Number
TC_PART_NAME  = "COMPONENT_NAME"   # TC Item Name
TC_OBJECT_NAME = "OBJECT_NAME"     # TC Object Name fallback
TC_OBJECT_DESC = "OBJECT_DESC"     # TC Description

# =============================================================================
# BOM DATA CLASS
# =============================================================================
@dataclass
class BOMItem:
    item_number : int   = 0
    part_number : str   = ""
    part_name   : str   = ""
    quantity    : int   = 0
    comp_tag    : NXOpen.Tag = NXOpen.Tag.Null

# =============================================================================
# MAIN
# =============================================================================
def main():

    mark_id = the_session.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible, "BOM_TC_NX2306_Python")

    # --- Validate ---
    if work_part is None:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "BOM Error", NXOpen.NXMessageBox.DialogType.Error,
            "No active part found.")
        return

    current_sheet = work_part.DrawingSheets.CurrentDrawingSheet
    if current_sheet is None:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "BOM Error", NXOpen.NXMessageBox.DialogType.Error,
            "No active drawing sheet found.\nPlease open a drawing sheet first.")
        return

    try:
        # Step 1: Collect components from TC assembly
        bom_items = collect_tc_components()

        if not bom_items:
            NXOpen.UI.GetUI().NXMessageBox.Show(
                "BOM", NXOpen.NXMessageBox.DialogType.Warning,
                "No components found in assembly.\n"
                "Please check assembly structure in Teamcenter.")
            return

        # Preview collected items
        preview = f"Components found: {len(bom_items)}\n"
        preview += "─" * 50 + "\n"
        preview += f"{'ITEM':<6} {'PART NO':<20} {'PART NAME':<25} {'QTY'}\n"
        preview += "─" * 50 + "\n"
        for b in bom_items:
            preview += (f"{b.item_number:<6} "
                        f"{b.part_number:<20} "
                        f"{b.part_name:<25} "
                        f"{b.quantity}\n")

        print(preview)  # Printed to NX output window
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "BOM Preview", NXOpen.NXMessageBox.DialogType.Information,
            preview + "\nCheck NX output window for details.")

        # Step 2: Create Part List Table (4 columns)
        create_part_list_4col(current_sheet, bom_items)

        # Step 3: Create Balloons
        create_balloons_tc(current_sheet, bom_items)

        NXOpen.UI.GetUI().NXMessageBox.Show(
            "BOM Complete", NXOpen.NXMessageBox.DialogType.Information,
            f"BOM Created Successfully!\n"
            f"Items : {len(bom_items)}\n"
            f"Sheet : {current_sheet.Name}")

    except Exception as ex:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "BOM Error", NXOpen.NXMessageBox.DialogType.Error,
            f"Error:\n{str(ex)}")

# =============================================================================
# STEP 1 - Collect Components from TC Integrated Assembly
# =============================================================================
def collect_tc_components() -> List[BOMItem]:

    bom_list  : List[BOMItem]    = []
    part_dict : Dict[str, int]   = {}   # part_number -> index in bom_list
    item_no   : int              = 1

    try:
        root_comp = work_part.ComponentAssembly.RootComponent

        # Single part (no assembly)
        if root_comp is None:
            single             = BOMItem()
            single.item_number = 1
            single.part_number = get_tc_part_number(work_part)
            single.part_name   = get_tc_part_name(work_part)
            single.quantity    = 1
            single.comp_tag    = NXOpen.Tag.Null
            bom_list.append(single)
            return bom_list

        # Traverse top-level children
        for child in root_comp.GetChildren():

            c_part = child.Prototype
            if not isinstance(c_part, NXOpen.Part):
                continue

            part = c_part  # type: NXOpen.Part

            # Get Part Number from TC
            p_no = get_tc_part_number(part)

            if p_no in part_dict:
                # Duplicate - increment quantity
                idx              = part_dict[p_no]
                bom_list[idx].quantity += 1
            else:
                # New part
                new_item             = BOMItem()
                new_item.item_number = item_no
                new_item.part_number = p_no
                new_item.part_name   = get_tc_part_name(part)
                new_item.quantity    = 1
                new_item.comp_tag    = child.Tag

                # Fallbacks
                if not new_item.part_name:
                    new_item.part_name = part.Name

                part_dict[p_no] = len(bom_list)
                bom_list.append(new_item)
                item_no += 1

    except Exception as ex:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "Error", NXOpen.NXMessageBox.DialogType.Error,
            f"Component collection error:\n{str(ex)}")

    return bom_list

# =============================================================================
# STEP 2 - Create Part List with 4 Columns (NX 2306 PartListBuilder)
# =============================================================================
def create_part_list_4col(sheet    : NXOpen.Drawings.DrawingSheet,
                           bom_items: List[BOMItem]):

    pl_builder = None

    try:
        pl_builder = work_part.Annotations.CreatePartListBuilder(None)

        # --- Settings ---
        pl_builder.Settings.Associative = True
        pl_builder.Settings.BomLevel    = (
            NXOpen.Annotations.PartListBuilder.PartListBomLevelOption.TopLevelOnly)
        pl_builder.Settings.SortingLevel = (
            NXOpen.Annotations.PartListBuilder.PartListSortingLevelOption.Level1)
        pl_builder.Settings.RowHeight    = 8.0

        # --- Position on sheet ---
        # Adjust X,Y to match your TC drawing title block position
        pl_builder.Origin.Anchor = (
            NXOpen.Annotations.OriginBuilder.LocationType.BottomRight)
        pl_builder.Origin.XValue = 400.0   # Adjust to your title block
        pl_builder.Origin.YValue =  10.0   # Adjust to your title block

        # --- Configure 4 Columns ---
        configure_4_columns(pl_builder)

        # --- Commit ---
        part_list = pl_builder.Commit()

        if part_list is not None:
            rows = part_list.GetAllMemberRows()
            NXOpen.UI.GetUI().NXMessageBox.Show(
                "Part List", NXOpen.NXMessageBox.DialogType.Information,
                f"Part List created with {len(rows)} rows.")
        else:
            NXOpen.UI.GetUI().NXMessageBox.Show(
                "Part List", NXOpen.NXMessageBox.DialogType.Warning,
                "Part List returned None.\nCheck assembly structure.")

    except Exception as ex:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "Part List Error", NXOpen.NXMessageBox.DialogType.Error,
            f"Part List creation error:\n{str(ex)}")
    finally:
        if pl_builder is not None:
            pl_builder.Destroy()

# =============================================================================
# 4 Column Configuration mapped to Teamcenter attributes
# =============================================================================
def configure_4_columns(pl_builder: NXOpen.Annotations.PartListBuilder):

    ColType = NXOpen.Annotations.PartListColumnBuilder.ColumnTypeOption
    Align   = NXOpen.Annotations.PartListColumnBuilder.ColumnAlignmentOption

    # ------------------------------------------------------------------
    # COLUMN 1 - ITEM NUMBER
    # FindNumber is auto-generated and linked to balloons
    # ------------------------------------------------------------------
    col1 = pl_builder.Columns.CreateColumnBuilder(None)
    col1.ColumnType = ColType.FindNumber
    col1.Heading.Text.SetText(["ITEM"])
    col1.Width     = 15.0
    col1.Alignment = Align.Center
    col1.Commit()

    # ------------------------------------------------------------------
    # COLUMN 2 - PART NUMBER (TC: DB_PART_NO = Item ID)
    # ------------------------------------------------------------------
    col2 = pl_builder.Columns.CreateColumnBuilder(None)
    col2.ColumnType    = ColType.Attribute
    col2.AttributeName = TC_ITEM_ID          # "DB_PART_NO"
    col2.Heading.Text.SetText(["PART NUMBER"])
    col2.Width         = 50.0
    col2.Alignment     = Align.Left
    col2.Commit()

    # ------------------------------------------------------------------
    # COLUMN 3 - QUANTITY (auto-counted from TC instances)
    # ------------------------------------------------------------------
    col3 = pl_builder.Columns.CreateColumnBuilder(None)
    col3.ColumnType = ColType.Quantity
    col3.Heading.Text.SetText(["QTY"])
    col3.Width     = 15.0
    col3.Alignment = Align.Center
    col3.Commit()

    # ------------------------------------------------------------------
    # COLUMN 4 - PART NAME (TC: COMPONENT_NAME)
    # ------------------------------------------------------------------
    col4 = pl_builder.Columns.CreateColumnBuilder(None)
    col4.ColumnType    = ColType.Attribute
    col4.AttributeName = TC_PART_NAME        # "COMPONENT_NAME"
    col4.Heading.Text.SetText(["PART NAME"])
    col4.Width         = 70.0
    col4.Alignment     = Align.Left
    col4.Commit()

# =============================================================================
# STEP 3 - Create Balloons (NX 2306 TC)
# =============================================================================
def create_balloons_tc(sheet    : NXOpen.Drawings.DrawingSheet,
                        bom_items: List[BOMItem]):

    # Get views on current sheet
    sheet_views = []
    for dv in work_part.DraftingViews:
        try:
            if dv.Sheet is not None and dv.Sheet.Tag == sheet.Tag:
                sheet_views.append(dv)
        except Exception:
            continue

    if not sheet_views:
        NXOpen.UI.GetUI().NXMessageBox.Show(
            "Balloon", NXOpen.NXMessageBox.DialogType.Warning,
            "No views found on sheet. Balloons skipped.")
        return

    # Try AutoBalloon first
    if try_auto_balloon():
        return

    # Fallback: Manual balloon placement
    manual_balloons_tc(sheet_views[0], bom_items)

# =============================================================================
# AutoBalloon for TC NX 2306
# =============================================================================
def try_auto_balloon() -> bool:

    try:
        ab_builder = work_part.Annotations.CreateAutoBalloonBuilder()

        ab_builder.BalloonType = (
            NXOpen.Annotations.AutoBalloonBuilder.BalloonShapeOption.CircularBalloon)

        ab_builder.Size = 8.0

        ab_builder.AttachmentType = (
            NXOpen.Annotations.AutoBalloonBuilder
            .AttachmentTypeOption.SilhouetteEdge)

        ab_builder.IgnoreExistingBalloons = True

        ab_builder.Commit()
        ab_builder.Destroy()

        NXOpen.UI.GetUI().NXMessageBox.Show(
            "Balloon", NXOpen.NXMessageBox.DialogType.Information,
            "Auto Balloons created and linked to Part List.")
        return True

    except Exception:
        return False

# =============================================================================
# Manual Balloon Placement for TC NX 2306
# =============================================================================
def manual_balloons_tc(main_view : NXOpen.Drawings.DraftingView,
                        bom_items : List[BOMItem]):

    count   = 0
    y_offset = 0.0

    for item in bom_items:
        try:
            # Get component center on sheet
            comp_pt = map_component_to_sheet(item.comp_tag, main_view, y_offset)

            # Offset balloon position
            sym_pt = NXOpen.Point3d(comp_pt.X + 25.0,
                                    comp_pt.Y + 10.0,
                                    0.0)

            # --- Create IdSymbol (Balloon) ---
            id_builder = work_part.Annotations.CreateIdSymbolBuilder(None)

            id_builder.Type      = (
                NXOpen.Annotations.IdSymbolBuilder.SymbolType.Circle)
            id_builder.UpperText = str(item.item_number)
            id_builder.LowerText = ""
            id_builder.Size      = 8.0

            # Position
            id_builder.Origin.Anchor = (
                NXOpen.Annotations.OriginBuilder.LocationType.AbsoluteXy)
            id_builder.Origin.XValue = sym_pt.X
            id_builder.Origin.YValue = sym_pt.Y

            # Leader line to component
            leader_pts    = [NXOpen.Point3d(comp_pt.X, comp_pt.Y, 0.0)]
            id_builder.Leader.Leaders.Item(0).SetLeaderPoints(leader_pts)

            # Associate balloon to component (TC NX 2306)
            if item.comp_tag != NXOpen.Tag.Null:
                try:
                    tag_obj = the_session.GetObjectManager() \
                                         .GetTaggedObject(item.comp_tag)
                    if isinstance(tag_obj, NXOpen.NXObject):
                        id_builder.AssociatedObject = tag_obj
                except Exception:
                    pass

            id_builder.Commit()
            id_builder.Destroy()

            count    += 1
            y_offset += 18.0

        except Exception:
            continue

    NXOpen.UI.GetUI().NXMessageBox.Show(
        "Balloon", NXOpen.NXMessageBox.DialogType.Information,
        f"Manual Balloons placed: {count}")

# =============================================================================
# HELPER - Get TC Part Number
# Priority: DB_PART_NO > OBJECT_NAME > UF Part Name > Part.Name
# =============================================================================
def get_tc_part_number(part: NXOpen.Part) -> str:

    val = safe_get_tc_attr(part, "DB_PART_NO")
    if val: return val

    val = safe_get_tc_attr(part, "OBJECT_NAME")
    if val: return val

    try:
        pname = the_uf_session.Part.AskPartName(part.Tag)
        if pname: return pname
    except Exception:
        pass

    return part.Name

# =============================================================================
# HELPER - Get TC Part Name
# Priority: COMPONENT_NAME > OBJECT_DESC > OBJECT_NAME > Part.Name
# =============================================================================
def get_tc_part_name(part: NXOpen.Part) -> str:

    val = safe_get_tc_attr(part, "COMPONENT_NAME")
    if val: return val

    val = safe_get_tc_attr(part, "OBJECT_DESC")
    if val: return val

    val = safe_get_tc_attr(part, "OBJECT_NAME")
    if val: return val

    return part.Name

# =============================================================================
# HELPER - Safe TC Attribute Reader (3 fallback methods)
# =============================================================================
def safe_get_tc_attr(part: NXOpen.Part, attr_name: str) -> str:

    # Method 1: GetUserAttribute (TC synced attributes)
    try:
        attr = part.GetUserAttribute(
            attr_name,
            NXOpen.NXObject.AttributeType.String, -1)
        if attr.StringValue and attr.StringValue.strip():
            return attr.StringValue.strip()
    except Exception:
        pass

    # Method 2: UF Attr read (legacy TC method)
    try:
        val = the_uf_session.Attr.ReadValueAsString(part.Tag, attr_name)
        if val and val.strip():
            return val.strip()
    except Exception:
        pass

    # Method 3: Scan all string attributes
    try:
        all_attrs = part.GetAllAttributesByType(
            NXOpen.NXObject.AttributeType.String)
        for a in all_attrs:
            if a.Title.upper().strip() == attr_name.upper().strip():
                if a.StringValue and a.StringValue.strip():
                    return a.StringValue.strip()
    except Exception:
        pass

    return ""

# =============================================================================
# HELPER - Map 3D component center to 2D sheet coordinates
# =============================================================================
def map_component_to_sheet(comp_tag : NXOpen.Tag,
                            d_view   : NXOpen.Drawings.DraftingView,
                            y_off    : float) -> NXOpen.Point3d:

    fallback = NXOpen.Point3d(
        d_view.Origin.X + 20.0,
        d_view.Origin.Y + 20.0 + y_off,
        0.0)

    try:
        if comp_tag != NXOpen.Tag.Null:

            mn, mx = the_uf_session.Modl.AskBoundingBoxExact(
                comp_tag, NXOpen.Tag.Null)

            center = [
                (mn[0] + mx[0]) / 2.0,
                (mn[1] + mx[1]) / 2.0,
                (mn[2] + mx[2]) / 2.0
            ]

            sheet_pt = the_uf_session.Draw.MapModelToDrawing(
                d_view.Tag, center)

            return NXOpen.Point3d(sheet_pt[0], sheet_pt[1], 0.0)

    except Exception:
        pass

    return fallback

# =============================================================================
# DIAGNOSTIC - List all TC attributes on work part
# Run this first to discover exact attribute names in your TC setup
# =============================================================================
def diagnose_tc_attributes():

    if work_part is None:
        return

    msg = f"TC Attributes on: {work_part.Name}\n"
    msg += "─" * 40 + "\n"

    try:
        all_attrs = work_part.GetAllAttributesByType(
            NXOpen.NXObject.AttributeType.String)

        if not all_attrs:
            msg += "No string attributes found.\n"
        else:
            for a in all_attrs:
                msg += f"{a.Title:<30} = {a.StringValue}\n"

    except Exception as ex:
        msg += f"Error: {str(ex)}"

    print(msg)   # Print to NX output window
    NXOpen.UI.GetUI().NXMessageBox.Show(
        "TC Diagnostic", NXOpen.NXMessageBox.DialogType.Information,
        msg[:2000])  # MsgBox truncated to 2000 chars, full list in output window

# =============================================================================
# UNLOAD
# =============================================================================
def get_unload_option(dummy: str) -> int:
    return NXOpen.Session.LibraryUnloadOption.Immediately

# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    main()
