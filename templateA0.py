# =============================================================================
# NX Drawing Template Customization Journal
# NXOpen Python API - Teamcenter Managed Mode Compatible
# =============================================================================
# Description:
#   Automates creation and update of ISO drawing templates in NX Drafting.
#   Supports A0–A4 sheet sizes, grid block generation, and Teamcenter
#   title block attribute mapping.
#
# Requirements:
#   - Siemens NX with NXOpen Python API
#   - Teamcenter managed mode (optional; gracefully degrades in unmanaged mode)
#
# Author  : NXOpen Python Developer
# Version : 1.0
# =============================================================================

import math
import NXOpen
import NXOpen.Drawings
import NXOpen.UF
import NXOpen.annotations

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------

# ISO sheet sizes in millimetres {name: (width_mm, height_mm)}
ISO_SHEET_SIZES = {
    "A0": (1189.0, 841.0),
    "A1": (841.0,  594.0),
    "A2": (594.0,  420.0),
    "A3": (420.0,  297.0),
    "A4": (297.0,  210.0),
}

# Grid block dimensions (mm)
HORIZ_BLOCK_SIZE = 30.0   # Each horizontal (column) block width
VERT_BLOCK_SIZE  = 24.0   # Each vertical (row) block height

# Sheet margin before grid starts (mm)
MARGIN = 10.0

# Title block reserved height at the bottom of the sheet (mm)
TITLE_BLOCK_HEIGHT = 40.0

# Default value when a Teamcenter attribute is missing
TC_DEFAULT = "-"

# Drawing line layer (put all template geometry on layer 200)
TEMPLATE_LAYER = 200


# ---------------------------------------------------------------------------
# HELPER UTILITIES
# ---------------------------------------------------------------------------

def mm_to_part_units(session: NXOpen.Session, mm_value: float) -> float:
    """
    Convert millimetres to the current part's unit system.
    NX internally works in part units; most metric parts use mm already,
    but we query the part unit scale to be safe.
    """
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    part = session.Parts.Work
    # UFPart.AskUnits returns the unit type: 1 = mm, 2 = inch
    unit_type = uf_session.Part.AskUnits(part.Tag)
    if unit_type == 2:          # inches
        return mm_value / 25.4
    return mm_value             # millimetres (default)


def get_or_create_layer_category(session: NXOpen.Session, layer: int):
    """Ensure the given layer is visible and selectable."""
    work_part = session.Parts.Work
    layer_state = work_part.Layers.GetState(layer)
    if layer_state == NXOpen.Layer.State.WorkLayer:
        return
    work_part.Layers.SetState(layer, NXOpen.Layer.State.Selectable)


def set_work_layer(session: NXOpen.Session, layer: int):
    """Set the active work layer."""
    work_part = session.Parts.Work
    work_part.Layers.WorkLayer = layer


# ---------------------------------------------------------------------------
# 1. SHEET CREATION / UPDATE
# ---------------------------------------------------------------------------

def create_sheet(session: NXOpen.Session,
                 sheet_name: str = "Template",
                 size_key: str  = "A3") -> NXOpen.Drawings.DrawingSheet:
    """
    Create a new drawing sheet or update an existing one.

    Parameters
    ----------
    session    : Active NXOpen session.
    sheet_name : Name for the drawing sheet.
    size_key   : ISO size string – "A0", "A1", "A2", "A3", or "A4".

    Returns
    -------
    The created / updated DrawingSheet object.
    """
    if size_key not in ISO_SHEET_SIZES:
        raise ValueError(f"Unknown sheet size '{size_key}'. "
                         f"Valid options: {list(ISO_SHEET_SIZES.keys())}")

    width_mm, height_mm = ISO_SHEET_SIZES[size_key]
    work_part = session.Parts.Work

    # -----------------------------------------------------------------------
    # Check whether the sheet already exists
    # -----------------------------------------------------------------------
    existing_sheet = None
    for sheet in work_part.DrawingSheets:
        if sheet.Name == sheet_name:
            existing_sheet = sheet
            break

    if existing_sheet is not None:
        # --- UPDATE existing sheet dimensions ---
        print(f"[create_sheet] Sheet '{sheet_name}' already exists – updating.")
        sheet_builder = work_part.DrawingSheets.CreateDrawingSheetBuilder(existing_sheet)
    else:
        # --- CREATE new sheet ---
        print(f"[create_sheet] Creating new sheet '{sheet_name}' ({size_key}).")
        sheet_builder = work_part.DrawingSheets.CreateDrawingSheetBuilder(
            NXOpen.NXObject.Null
        )
        sheet_builder.SheetName = sheet_name

    # Apply dimensions (NX sheet builder uses mm directly in metric parts)
    sheet_builder.Width  = width_mm
    sheet_builder.Height = height_mm

    # Standard ISO scale 1:1 (stored as ratio denominator)
    sheet_builder.Scale = 1.0

    # Projection angle – First angle (ISO default = 1; Third angle = 3)
    sheet_builder.ProjectionAngle = (
        NXOpen.Drawings.DrawingSheetBuilder.ProjectionAngleType.FirstAngle
    )

    # Commit the builder
    nxobject = sheet_builder.Commit()
    sheet_builder.Destroy()

    # Retrieve the DrawingSheet object
    result_sheet = None
    for sheet in work_part.DrawingSheets:
        if sheet.Name == sheet_name:
            result_sheet = sheet
            break

    if result_sheet is None:
        raise RuntimeError("[create_sheet] Failed to locate sheet after creation.")

    # Make the new sheet the display sheet
    result_sheet.Open()

    print(f"[create_sheet] Sheet '{sheet_name}' ready "
          f"({width_mm} x {height_mm} mm).")
    return result_sheet


# ---------------------------------------------------------------------------
# 2. GRID BLOCK CREATION
# ---------------------------------------------------------------------------

def _delete_existing_grid(work_part: NXOpen.Part, layer: int):
    """
    Remove all drafting lines/notes previously placed on the template layer
    so the grid can be regenerated cleanly after a sheet-size change.
    """
    objects_to_delete = []
    for obj in work_part.Lines:
        if obj.Layer == layer:
            objects_to_delete.append(obj)
    for obj in work_part.Annotations.GetAnnotations():
        if obj.Layer == layer:
            objects_to_delete.append(obj)

    if objects_to_delete:
        mark_id = NXOpen.Session.GetSession().SetUndoMark(
            NXOpen.Session.MarkVisibility.Invisible, "DeleteOldGrid"
        )
        NXOpen.Session.GetSession().UpdateManager.AddObjectsToDeleteList(
            objects_to_delete
        )
        NXOpen.Session.GetSession().UpdateManager.DoUpdate(mark_id)
        print(f"[_delete_existing_grid] Removed {len(objects_to_delete)} "
              "old grid objects.")


def _draw_line(work_part: NXOpen.Part,
               x1: float, y1: float,
               x2: float, y2: float,
               layer: int) -> NXOpen.Line:
    """
    Create a drafting line between two 2-D points (Z = 0) on the given layer.
    """
    start = NXOpen.Point3d(x1, y1, 0.0)
    end   = NXOpen.Point3d(x2, y2, 0.0)
    line  = work_part.Curves.CreateLine(start, end)
    line.Layer = layer
    return line


def _place_label(work_part: NXOpen.Part,
                 text: str,
                 x: float, y: float,
                 layer: int,
                 height_mm: float = 3.5):
    """
    Place a centred drafting note at position (x, y).

    Parameters
    ----------
    text      : Label string.
    x, y      : Centre point of the label in sheet coordinates (mm).
    layer     : NX layer number.
    height_mm : Text character height in mm.
    """
    note_builder = work_part.Annotations.CreateDraftingNoteBuilder(
        NXOpen.Annotations.DraftingNote.Null
    )

    # Text content
    note_builder.Text.TextBlock.SetText([text])

    # Formatting
    note_builder.Style.LetteringStyle.GeneralTextSize = height_mm
    note_builder.Style.LetteringStyle.AnchorAlignment = (
        NXOpen.Annotations.LetteringPreferences.Alignment.Center
    )

    # Origin
    note_builder.Origin.SetInferRelativeToGeometry(False)
    note_builder.Origin.Anchor = (
        NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
    )
    origin_pt = NXOpen.Point3d(x, y, 0.0)
    note_builder.Origin.SetOriginPoint(origin_pt)

    note = note_builder.Commit()
    note_builder.Destroy()
    note.Layer = layer
    return note


def create_grid(session: NXOpen.Session,
                sheet: NXOpen.Drawings.DrawingSheet,
                size_key: str = "A3"):
    """
    Generate the ISO-style grid (column numbers + row letters) around the
    usable drawing area.

    Layout
    ------
    - A margin of MARGIN mm is kept on all four sides.
    - An additional TITLE_BLOCK_HEIGHT mm is reserved at the bottom for the
      title block (grid does not cover title block area).
    - Columns (numeric labels 1, 2, 3 …) run left → right.
    - Rows    (alpha  labels A, B, C …) run top  → bottom.
    - Border lines and inner tick lines are drawn on TEMPLATE_LAYER.
    - Labels are centred inside each border cell.
    """
    if size_key not in ISO_SHEET_SIZES:
        raise ValueError(f"[create_grid] Unknown size key '{size_key}'.")

    width_mm, height_mm = ISO_SHEET_SIZES[size_key]
    work_part = session.Parts.Work

    # Remove old grid objects to avoid duplication on re-run
    _delete_existing_grid(work_part, TEMPLATE_LAYER)

    # Ensure template layer is accessible
    get_or_create_layer_category(session, TEMPLATE_LAYER)
    set_work_layer(session, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Geometry boundaries
    # -----------------------------------------------------------------------
    # Sheet origin is (0, 0) bottom-left in NX drawing coordinates.

    # Outer border of the drawing frame (inside the sheet margin)
    border_left   = MARGIN
    border_right  = width_mm  - MARGIN
    border_bottom = MARGIN + TITLE_BLOCK_HEIGHT   # title block sits here
    border_top    = height_mm - MARGIN

    usable_width  = border_right  - border_left
    usable_height = border_top    - border_bottom

    # Number of full blocks that fit
    num_cols = int(usable_width  / HORIZ_BLOCK_SIZE)
    num_rows = int(usable_height / VERT_BLOCK_SIZE)

    if num_cols < 1 or num_rows < 1:
        raise RuntimeError("[create_grid] Sheet too small to fit any grid block.")

    # Actual grid extents (centred within the usable area if blocks don't
    # fill it exactly)
    grid_width  = num_cols * HORIZ_BLOCK_SIZE
    grid_height = num_rows * VERT_BLOCK_SIZE

    grid_left   = border_left  + (usable_width  - grid_width)  / 2.0
    grid_right  = grid_left    + grid_width
    grid_bottom = border_bottom + (usable_height - grid_height) / 2.0
    grid_top    = grid_bottom  + grid_height

    # Width of the border label cells (same as one block)
    cell_w = HORIZ_BLOCK_SIZE
    cell_h = VERT_BLOCK_SIZE

    print(f"[create_grid] Sheet {size_key}: "
          f"{num_cols} columns × {num_rows} rows  "
          f"(grid {grid_width:.1f} × {grid_height:.1f} mm)")

    # -----------------------------------------------------------------------
    # Draw outer rectangle of the drawing frame
    # -----------------------------------------------------------------------
    _draw_line(work_part, border_left,  border_bottom, border_right, border_bottom, TEMPLATE_LAYER)
    _draw_line(work_part, border_right, border_bottom, border_right, border_top,    TEMPLATE_LAYER)
    _draw_line(work_part, border_right, border_top,    border_left,  border_top,    TEMPLATE_LAYER)
    _draw_line(work_part, border_left,  border_top,    border_left,  border_bottom, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Draw vertical grid lines (column dividers) + top/bottom tick borders
    # -----------------------------------------------------------------------
    for col in range(num_cols + 1):
        x = grid_left + col * cell_w

        # Full-height inner grid line
        _draw_line(work_part, x, grid_bottom, x, grid_top, TEMPLATE_LAYER)

        # Top border cell line  (between grid_top and border_top)
        _draw_line(work_part, x, grid_top,    x, border_top,    TEMPLATE_LAYER)

        # Bottom border cell line (between border_bottom and grid_bottom)
        _draw_line(work_part, x, border_bottom, x, grid_bottom, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Draw horizontal grid lines (row dividers) + left/right tick borders
    # -----------------------------------------------------------------------
    for row in range(num_rows + 1):
        y = grid_bottom + row * cell_h

        # Full-width inner grid line
        _draw_line(work_part, grid_left, y, grid_right, y, TEMPLATE_LAYER)

        # Left border cell line
        _draw_line(work_part, border_left, y, grid_left, y, TEMPLATE_LAYER)

        # Right border cell line
        _draw_line(work_part, grid_right, y, border_right, y, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Horizontal border (top / bottom) label lines
    # -----------------------------------------------------------------------
    # Top horizontal separator between border band and grid area
    _draw_line(work_part, border_left, grid_top, border_right, grid_top, TEMPLATE_LAYER)

    # Bottom horizontal separator between grid area and border band
    _draw_line(work_part, border_left, grid_bottom, border_right, grid_bottom, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Vertical border (left / right) separator lines
    # -----------------------------------------------------------------------
    _draw_line(work_part, grid_left,  border_bottom, grid_left,  border_top, TEMPLATE_LAYER)
    _draw_line(work_part, grid_right, border_bottom, grid_right, border_top, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Column numeric labels  (1, 2, 3 … num_cols)
    # Placed in the top AND bottom border bands, centred in each column cell.
    # -----------------------------------------------------------------------
    for col in range(num_cols):
        label = str(col + 1)
        cx = grid_left + col * cell_w + cell_w / 2.0

        # Top border band
        cy_top = grid_top + (border_top - grid_top) / 2.0
        _place_label(work_part, label, cx, cy_top, TEMPLATE_LAYER)

        # Bottom border band
        cy_bot = border_bottom + (grid_bottom - border_bottom) / 2.0
        _place_label(work_part, label, cx, cy_bot, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Row alphabetic labels  (A, B, C … )
    # Placed in the left AND right border bands, centred in each row cell.
    # Rows are numbered top → bottom so row 0 = top row.
    # -----------------------------------------------------------------------
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for row in range(num_rows):
        # Clamp label to alphabet; if more than 26 rows wrap with AA, AB, …
        if row < 26:
            label = alphabet[row]
        else:
            label = alphabet[row // 26 - 1] + alphabet[row % 26]

        # Row 0 is at the top, so invert y
        cy = grid_top - row * cell_h - cell_h / 2.0

        # Left border band
        cx_left = border_left + (grid_left - border_left) / 2.0
        _place_label(work_part, label, cx_left, cy, TEMPLATE_LAYER)

        # Right border band
        cx_right = grid_right + (border_right - grid_right) / 2.0
        _place_label(work_part, label, cx_right, cy, TEMPLATE_LAYER)

    print(f"[create_grid] Grid created: {num_cols} columns, {num_rows} rows.")


# ---------------------------------------------------------------------------
# 3. TEAMCENTER ATTRIBUTE RETRIEVAL
# ---------------------------------------------------------------------------

def get_tc_attributes(session: NXOpen.Session) -> dict:
    """
    Read Teamcenter item attributes from the work part's TC component.

    Attempts to read the following standard TC attributes:
        item_id       – Item ID
        item_revision – Revision ID
        object_name   – Object / Part name
        drawing_no    – Drawing number (custom attribute)
        material      – Material (custom attribute)

    Returns a dict.  Missing attributes default to TC_DEFAULT ("-").
    Falls back gracefully when running in unmanaged (non-TC) mode.
    """
    attrs = {
        "item_id":       TC_DEFAULT,
        "item_revision": TC_DEFAULT,
        "object_name":   TC_DEFAULT,
        "drawing_no":    TC_DEFAULT,
        "material":      TC_DEFAULT,
    }

    try:
        uf_session = NXOpen.UF.UFSession.GetUFSession()
        work_part  = session.Parts.Work

        # ------------------------------------------------------------------
        # Check managed mode
        # ------------------------------------------------------------------
        pdm_session = uf_session.Pdm
        is_managed  = pdm_session.IsPartManaged(work_part.Tag)

        if not is_managed:
            print("[get_tc_attributes] Part is NOT in managed mode. "
                  "Using default attribute values.")
            return attrs

        # ------------------------------------------------------------------
        # Retrieve TC object tag for the work part
        # ------------------------------------------------------------------
        # UFPdm.AskWorkPartItem returns (item_tag, item_revision_tag)
        item_tag, item_rev_tag = pdm_session.AskWorkPartItem(work_part.Tag)

        # ------------------------------------------------------------------
        # Item ID
        # ------------------------------------------------------------------
        try:
            item_id_val = uf_session.Obj.AskStringAttr(item_tag, "item_id")
            attrs["item_id"] = item_id_val if item_id_val else TC_DEFAULT
        except Exception:
            pass

        # ------------------------------------------------------------------
        # Item Revision
        # ------------------------------------------------------------------
        try:
            rev_val = uf_session.Obj.AskStringAttr(item_rev_tag, "item_revision_id")
            attrs["item_revision"] = rev_val if rev_val else TC_DEFAULT
        except Exception:
            pass

        # ------------------------------------------------------------------
        # Object Name
        # ------------------------------------------------------------------
        try:
            name_val = uf_session.Obj.AskName(item_tag)
            attrs["object_name"] = name_val if name_val else TC_DEFAULT
        except Exception:
            pass

        # ------------------------------------------------------------------
        # Custom attributes via UFAttr (drawing number, material)
        # ------------------------------------------------------------------
        custom_map = {
            "drawing_no": ["drawing_number", "DrawingNumber", "DRW_NO"],
            "material":   ["material",        "Material",      "MAT"],
        }

        for key, candidates in custom_map.items():
            for attr_name in candidates:
                try:
                    # Try on item revision first, then item
                    for tag in (item_rev_tag, item_tag):
                        val = uf_session.Obj.AskStringAttr(tag, attr_name)
                        if val:
                            attrs[key] = val
                            raise StopIteration   # break both loops
                except StopIteration:
                    break
                except Exception:
                    continue

        print(f"[get_tc_attributes] Retrieved: {attrs}")

    except Exception as ex:
        print(f"[get_tc_attributes] Warning – could not read TC attributes: {ex}")

    return attrs


# ---------------------------------------------------------------------------
# 4. TITLE BLOCK UPDATE
# ---------------------------------------------------------------------------

def _find_note_by_hint(work_part: NXOpen.Part, hint: str):
    """
    Search all drafting notes for one whose text starts with the given hint.
    Returns the first match or None.

    Hint format used in this template:
        "ITEM_ID:"  "REV:"  "NAME:"  "DWG:"  "MAT:"
    """
    for ann in work_part.Annotations.GetAnnotations():
        if not isinstance(ann, NXOpen.Annotations.DraftingNote):
            continue
        try:
            lines = ann.GetText()
            if lines and lines[0].startswith(hint):
                return ann
        except Exception:
            continue
    return None


def _set_note_value(note: NXOpen.Annotations.DraftingNote, value: str):
    """Update the second part of a 'KEY: value' drafting note."""
    try:
        lines = note.GetText()
        if lines:
            key_part = lines[0].split(":")[0]
            note.SetText([f"{key_part}: {value}"])
    except Exception as ex:
        print(f"[_set_note_value] Could not update note: {ex}")


def update_title_block(session: NXOpen.Session,
                       sheet: NXOpen.Drawings.DrawingSheet,
                       size_key: str = "A3"):
    """
    Draw (or update) the title block at the bottom of the drawing sheet
    and populate it with Teamcenter attributes.

    Title block layout (bottom TITLE_BLOCK_HEIGHT mm of the sheet):

    ┌─────────────────────────────────────────────┐
    │ ITEM ID:   <value>   │  DWG NO:   <value>   │
    │ NAME:      <value>   │  REV:      <value>   │
    │ MATERIAL:  <value>   │  SCALE:    1:1        │
    └─────────────────────────────────────────────┘
    """
    width_mm, height_mm = ISO_SHEET_SIZES[size_key]
    work_part = session.Parts.Work

    # Ensure template layer is set
    get_or_create_layer_category(session, TEMPLATE_LAYER)
    set_work_layer(session, TEMPLATE_LAYER)

    # Fetch TC attributes
    tc_attrs = get_tc_attributes(session)

    # -----------------------------------------------------------------------
    # Title block geometry coordinates
    # -----------------------------------------------------------------------
    tb_left   = MARGIN
    tb_right  = width_mm - MARGIN
    tb_bottom = MARGIN
    tb_top    = MARGIN + TITLE_BLOCK_HEIGHT

    tb_width  = tb_right - tb_left
    mid_x     = tb_left + tb_width / 2.0

    row_h     = TITLE_BLOCK_HEIGHT / 3.0    # three rows in title block

    # -----------------------------------------------------------------------
    # Title block outer border
    # -----------------------------------------------------------------------
    _draw_line(work_part, tb_left,  tb_bottom, tb_right, tb_bottom, TEMPLATE_LAYER)
    _draw_line(work_part, tb_right, tb_bottom, tb_right, tb_top,    TEMPLATE_LAYER)
    _draw_line(work_part, tb_right, tb_top,    tb_left,  tb_top,    TEMPLATE_LAYER)
    _draw_line(work_part, tb_left,  tb_top,    tb_left,  tb_bottom, TEMPLATE_LAYER)

    # Vertical centre divider
    _draw_line(work_part, mid_x, tb_bottom, mid_x, tb_top, TEMPLATE_LAYER)

    # Two horizontal row dividers
    for r in range(1, 3):
        y = tb_bottom + r * row_h
        _draw_line(work_part, tb_left, y, tb_right, y, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Helper to add/update a label note at (x, y)
    # If the note already exists (identified by its hint prefix) update it;
    # otherwise create a new one.
    # -----------------------------------------------------------------------
    def place_or_update(hint: str, value: str, x: float, y: float):
        full_text = f"{hint} {value}"
        existing  = _find_note_by_hint(work_part, hint)
        if existing is not None:
            existing.SetText([full_text])
            existing.Layer = TEMPLATE_LAYER
        else:
            _place_label(work_part, full_text, x, y,
                         TEMPLATE_LAYER, height_mm=3.0)

    # -----------------------------------------------------------------------
    # Row centres
    # -----------------------------------------------------------------------
    left_cx  = tb_left + tb_width / 4.0          # centre of left column
    right_cx = tb_left + 3.0 * tb_width / 4.0    # centre of right column

    # Row 1 (top row)  y-centre
    y_row1 = tb_bottom + 2.5 * row_h
    # Row 2 (middle)
    y_row2 = tb_bottom + 1.5 * row_h
    # Row 3 (bottom)
    y_row3 = tb_bottom + 0.5 * row_h

    # -----------------------------------------------------------------------
    # Place title block fields
    # Row 1: Item ID  |  Drawing No
    # Row 2: Name     |  Revision
    # Row 3: Material |  Scale
    # -----------------------------------------------------------------------
    place_or_update("ITEM ID:",   tc_attrs["item_id"],       left_cx,  y_row1)
    place_or_update("DWG NO:",    tc_attrs["drawing_no"],    right_cx, y_row1)

    place_or_update("NAME:",      tc_attrs["object_name"],   left_cx,  y_row2)
    place_or_update("REV:",       tc_attrs["item_revision"], right_cx, y_row2)

    place_or_update("MATERIAL:",  tc_attrs["material"],      left_cx,  y_row3)
    place_or_update("SCALE:",     "1:1",                     right_cx, y_row3)

    print("[update_title_block] Title block created/updated.")


# ---------------------------------------------------------------------------
# 5. MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def main():
    """
    Main journal entry point.

    Workflow
    --------
    1. Acquire the NX session and verify we are in the Drafting application.
    2. Prompt (via listing window) for the desired sheet size.
    3. Create or update the drawing sheet.
    4. Generate the grid.
    5. Draw / update the title block.
    6. Update the display.

    To change the target sheet size edit SELECTED_SIZE below or wire it up
    to an NX block-UI dialog for interactive selection.
    """
    # -----------------------------------------------------------------------
    # Acquire session
    # -----------------------------------------------------------------------
    session = NXOpen.Session.GetSession()

    # -----------------------------------------------------------------------
    # Configuration – change this to select a different sheet size
    # -----------------------------------------------------------------------
    SELECTED_SIZE  = "A3"          # Options: "A0", "A1", "A2", "A3", "A4"
    SHEET_NAME     = "DrawingSheet1"

    # -----------------------------------------------------------------------
    # Validate environment
    # -----------------------------------------------------------------------
    work_part = session.Parts.Work
    if work_part is None:
        session.ListingWindow.Open()
        session.ListingWindow.WriteLine(
            "[main] ERROR: No work part is open. Please open a part first."
        )
        return

    # Check that a drawing sheet context is accessible
    # (NX must be in Drafting or Gateway with a drawing part open)
    if not hasattr(work_part, "DrawingSheets"):
        session.ListingWindow.Open()
        session.ListingWindow.WriteLine(
            "[main] ERROR: Work part does not support drawing sheets."
        )
        return

    session.ListingWindow.Open()
    session.ListingWindow.WriteLine(
        "=" * 60 + "\n"
        "  NX Drawing Template Customization Journal\n"
        "=" * 60
    )
    session.ListingWindow.WriteLine(
        f"  Sheet Size : {SELECTED_SIZE}  "
        f"({ISO_SHEET_SIZES[SELECTED_SIZE][0]} x "
        f"{ISO_SHEET_SIZES[SELECTED_SIZE][1]} mm)"
    )
    session.ListingWindow.WriteLine(f"  Sheet Name : {SHEET_NAME}")
    session.ListingWindow.WriteLine("-" * 60)

    # -----------------------------------------------------------------------
    # Undo mark – allows single-step undo of the entire journal
    # -----------------------------------------------------------------------
    undo_mark = session.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible,
        "NX Drawing Template"
    )

    try:
        # -------------------------------------------------------------------
        # Step 1: Sheet
        # -------------------------------------------------------------------
        session.ListingWindow.WriteLine("[1/3] Creating / updating sheet …")
        sheet = create_sheet(session, SHEET_NAME, SELECTED_SIZE)

        # -------------------------------------------------------------------
        # Step 2: Grid
        # -------------------------------------------------------------------
        session.ListingWindow.WriteLine("[2/3] Building grid blocks …")
        create_grid(session, sheet, SELECTED_SIZE)

        # -------------------------------------------------------------------
        # Step 3: Title block
        # -------------------------------------------------------------------
        session.ListingWindow.WriteLine("[3/3] Updating title block …")
        update_title_block(session, sheet, SELECTED_SIZE)

        # -------------------------------------------------------------------
        # Refresh display
        # -------------------------------------------------------------------
        session.Parts.Work.Views.WorkView.UpdateDisplay()

        session.ListingWindow.WriteLine(
            "\n[main] Template customization complete.\n" + "=" * 60
        )

    except Exception as ex:
        # Undo everything on failure
        session.UndoToMark(undo_mark, "NX Drawing Template")
        session.ListingWindow.WriteLine(
            f"\n[main] FATAL ERROR – changes rolled back.\n  {ex}"
        )
        raise

    finally:
        # Clean up the undo mark (keep it so user can undo manually)
        session.DeleteUndoMark(undo_mark, "NX Drawing Template")


# ---------------------------------------------------------------------------
# NX journal execution entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
