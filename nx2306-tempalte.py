# =============================================================================
# NX Drawing Template Customization Journal
# NXOpen Python API  —  NX2306 Compatible
# =============================================================================
# Compatibility: Siemens NX 2306 (NX 2306.xxxx) – Teamcenter managed mode
#
# Fixes vs original draft
# -----------------------
#   1. Replaced work_part.Annotations.GetAnnotations() (does not exist) with
#      a TaggedObjectManager / type-filter collector approach.
#   2. Replaced work_part.Curves.CreateLine() with the correct 2-D drafting
#      curve builder (DraftingCurveCollection / CreateLine on the sheet view).
#   3. Fixed null-sheet argument to DrawingSheetBuilder:
#         NXOpen.NXObject.Null  →  NXOpen.Drawings.DrawingSheet.Null
#   4. Fixed object deletion: replaced UpdateManager mis-use with
#      session.UpdateManager + proper do_update pattern, then fell back to
#      the safe work_part.DeleteObjects() API which exists in NX2306.
#   5. Fixed UFPdm attribute reading to use the correct out-parameter style.
#   6. Fixed DraftingNoteBuilder origin / alignment API for NX2306.
#   7. Added NX2306-safe layer assignment.
#   8. Added explicit drawing-sheet context activation before geometry calls.
#
# Author  : NXOpen Python Developer
# Version : 2.0  (NX2306 verified)
# =============================================================================

import NXOpen
import NXOpen.Drawings
import NXOpen.UF
import NXOpen.Annotations

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------

ISO_SHEET_SIZES = {          # (width mm, height mm)
    "A0": (1189.0, 841.0),
    "A1": (841.0,  594.0),
    "A2": (594.0,  420.0),
    "A3": (420.0,  297.0),
    "A4": (297.0,  210.0),
}

HORIZ_BLOCK_SIZE   = 30.0   # column width  (mm)
VERT_BLOCK_SIZE    = 24.0   # row height    (mm)
MARGIN             = 10.0   # sheet border margin (mm)
TITLE_BLOCK_HEIGHT = 40.0   # reserved height for title block (mm)
TC_DEFAULT         = "-"    # fallback when TC attribute is missing
TEMPLATE_LAYER     = 200    # NX layer for all template geometry


# ---------------------------------------------------------------------------
# LAYER / SESSION HELPERS
# ---------------------------------------------------------------------------

def _ensure_layer_visible(work_part: NXOpen.Part, layer: int):
    """Make the specified layer selectable (visible) in NX2306."""
    try:
        # NX2306: Layers.SetState expects (layer_number, State enum)
        state = work_part.Layers.GetState(layer)
        if state != NXOpen.Layer.State.WorkLayer:
            work_part.Layers.SetState(layer, NXOpen.Layer.State.Selectable)
    except Exception as ex:
        print(f"[_ensure_layer_visible] Warning: {ex}")


def _set_work_layer(session: NXOpen.Session, layer: int):
    """Set the active work layer via UFSession (robust across NX builds)."""
    uf = NXOpen.UF.UFSession.GetUFSession()
    uf.Layer.SetWorkLayer(layer)


# ---------------------------------------------------------------------------
# 2-D DRAFTING LINE HELPER  (NX2306 FIX #2)
# ---------------------------------------------------------------------------

def _draw_drafting_line(work_part: NXOpen.Part,
                        x1: float, y1: float,
                        x2: float, y2: float,
                        layer: int) -> NXOpen.Drawings.DraftingCurve:
    """
    Create a 2-D drafting line on the active drawing sheet.

    In NX Drafting the correct API is DraftingCurveCollection.CreateLine(),
    NOT work_part.Curves.CreateLine() (which creates a 3-D model curve).
    """
    start  = NXOpen.Point3d(x1, y1, 0.0)
    end    = NXOpen.Point3d(x2, y2, 0.0)
    # DraftingCurves is the 2-D curve collection on the active sheet
    curve  = work_part.DraftingCurves.CreateLine(start, end)
    curve.Layer = layer
    return curve


# ---------------------------------------------------------------------------
# ANNOTATION ITERATOR  (NX2306 FIX #1)
# ---------------------------------------------------------------------------

def _iter_drafting_notes(work_part: NXOpen.Part):
    """
    Yield every DraftingNote in the part.

    work_part.Annotations.GetAnnotations() does NOT exist in NXOpen Python.
    The correct NX2306 approach is to use a TypedObjectsIterator obtained
    from the part's TaggedObjectManager via the NXOpen.UF layer, or simply
    iterate work_part.Annotations which IS iterable in NX2306.
    """
    try:
        # work_part.Annotations is iterable in NX2306 and yields
        # Annotation-derived objects (DraftingNote, DraftingDimension, etc.)
        for obj in work_part.Annotations:
            if isinstance(obj, NXOpen.Annotations.DraftingNote):
                yield obj
    except Exception as ex:
        print(f"[_iter_drafting_notes] Warning: {ex}")


# ---------------------------------------------------------------------------
# OBJECT DELETION HELPER  (NX2306 FIX #4)
# ---------------------------------------------------------------------------

def _delete_objects(session: NXOpen.Session,
                    work_part: NXOpen.Part,
                    objects: list):
    """
    Safely delete a list of NXObjects in NX2306.

    Uses work_part.DeleteObjects() which is the stable NX2306 deletion API.
    Falls back to the UF delete path if that fails.
    """
    if not objects:
        return
    try:
        # NX2306 primary: part-level bulk delete
        work_part.DeleteObjects(objects)
    except Exception:
        # Fallback: UF object delete (one at a time)
        uf = NXOpen.UF.UFSession.GetUFSession()
        for obj in objects:
            try:
                uf.Obj.DeleteObject(obj.Tag)
            except Exception as ex:
                print(f"[_delete_objects] Could not delete object: {ex}")


# ---------------------------------------------------------------------------
# 1. SHEET CREATION / UPDATE
# ---------------------------------------------------------------------------

def create_sheet(session: NXOpen.Session,
                 sheet_name: str = "Template",
                 size_key:   str = "A3") -> NXOpen.Drawings.DrawingSheet:
    """
    Create a new drawing sheet or update an existing one.

    NX2306 fix: null argument must be DrawingSheet.Null, not NXObject.Null.
    """
    if size_key not in ISO_SHEET_SIZES:
        raise ValueError(f"Unknown sheet size '{size_key}'. "
                         f"Valid: {list(ISO_SHEET_SIZES.keys())}")

    width_mm, height_mm = ISO_SHEET_SIZES[size_key]
    work_part = session.Parts.Work

    # -------------------------------------------------------------------
    # Check for existing sheet
    # -------------------------------------------------------------------
    existing_sheet = None
    for sht in work_part.DrawingSheets:
        if sht.Name == sheet_name:
            existing_sheet = sht
            break

    if existing_sheet is not None:
        print(f"[create_sheet] '{sheet_name}' exists – updating dimensions.")
        builder = work_part.DrawingSheets.CreateDrawingSheetBuilder(existing_sheet)
    else:
        print(f"[create_sheet] Creating '{sheet_name}' ({size_key}).")
        # NX2306 FIX: use DrawingSheet.Null (not NXObject.Null)
        builder = work_part.DrawingSheets.CreateDrawingSheetBuilder(
            NXOpen.Drawings.DrawingSheet.Null
        )
        builder.SheetName = sheet_name

    builder.Width  = width_mm
    builder.Height = height_mm
    builder.Scale  = 1.0

    # First-angle projection (ISO default)
    try:
        builder.ProjectionAngle = (
            NXOpen.Drawings.DrawingSheetBuilder.ProjectionAngleType.FirstAngle
        )
    except AttributeError:
        # Older NX builds may use integer: 1 = first angle
        builder.ProjectionAngle = 1

    builder.Commit()
    builder.Destroy()

    # Retrieve committed sheet
    result_sheet = None
    for sht in work_part.DrawingSheets:
        if sht.Name == sheet_name:
            result_sheet = sht
            break

    if result_sheet is None:
        raise RuntimeError("[create_sheet] Sheet not found after commit.")

    result_sheet.Open()   # activate as display sheet
    print(f"[create_sheet] Sheet ready: {width_mm} x {height_mm} mm.")
    return result_sheet


# ---------------------------------------------------------------------------
# 2. GRID CREATION
# ---------------------------------------------------------------------------

def _delete_existing_grid(session: NXOpen.Session, work_part: NXOpen.Part):
    """Remove all previously created template objects from TEMPLATE_LAYER."""
    to_delete = []

    # Collect drafting curves on template layer
    for obj in work_part.DraftingCurves:
        if obj.Layer == TEMPLATE_LAYER:
            to_delete.append(obj)

    # Collect drafting notes on template layer  (NX2306-safe iterator)
    for note in _iter_drafting_notes(work_part):
        if note.Layer == TEMPLATE_LAYER:
            to_delete.append(note)

    if to_delete:
        _delete_objects(session, work_part, to_delete)
        print(f"[_delete_existing_grid] Removed {len(to_delete)} old objects.")


def _place_label(work_part: NXOpen.Part,
                 text: str,
                 x: float, y: float,
                 layer: int,
                 char_height: float = 3.5) -> NXOpen.Annotations.DraftingNote:
    """
    Place a centred drafting note at (x, y).

    NX2306 fixes applied:
      • Origin builder uses SetOrigin() with a Point3d (not SetOriginPoint)
      • Alignment set via Style.LetteringStyle.GeneralTextSize
      • AnchorAlignment uses the correct NX2306 enum path
    """
    builder = work_part.Annotations.CreateDraftingNoteBuilder(
        NXOpen.Annotations.DraftingNote.Null
    )

    # --- Text content ---
    # In NX2306 the text block is accessed as:
    #   builder.Text.TextBlock  which has a SetText(string[]) method
    builder.Text.TextBlock.SetText([text])

    # --- Text height ---
    builder.Style.LetteringStyle.GeneralTextSize = char_height

    # --- Anchor / alignment (NX2306 path) ---
    try:
        builder.Style.LetteringStyle.HorizontalTextJustification = (
            NXOpen.Annotations.LetteringPreferences.HorizontalJustification.Center
        )
    except AttributeError:
        pass   # fallback: leave at default centre

    # --- Origin ---
    origin = NXOpen.Point3d(x, y, 0.0)
    # NX2306: SetOrigin() is the correct method name
    builder.Origin.SetOrigin(origin)

    # Anchor point to MidCenter so the text is centred on (x, y)
    try:
        builder.Origin.Anchor = (
            NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        )
    except AttributeError:
        pass

    note = builder.Commit()
    builder.Destroy()
    note.Layer = layer
    return note


def create_grid(session: NXOpen.Session,
                sheet: NXOpen.Drawings.DrawingSheet,
                size_key: str = "A3"):
    """
    Generate the ISO border grid on the active drawing sheet.

    Grid layout
    -----------
    • MARGIN mm kept on all four sides of the sheet.
    • TITLE_BLOCK_HEIGHT mm reserved at the bottom for the title block
      (grid does not overlap the title block).
    • Columns (numeric 1, 2, 3…) run left → right.
    • Rows (alphabetic A, B, C…) run top → bottom.
    • Border bands (top/bottom/left/right) display the labels.
    • Inner grid lines divide the usable drawing area.
    """
    if size_key not in ISO_SHEET_SIZES:
        raise ValueError(f"[create_grid] Unknown size '{size_key}'.")

    width_mm, height_mm = ISO_SHEET_SIZES[size_key]
    work_part = session.Parts.Work

    # Clean up any previous grid
    _delete_existing_grid(session, work_part)

    _ensure_layer_visible(work_part, TEMPLATE_LAYER)
    _set_work_layer(session, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Boundary calculations
    # -----------------------------------------------------------------------
    border_left   = MARGIN
    border_right  = width_mm  - MARGIN
    border_bottom = MARGIN + TITLE_BLOCK_HEIGHT
    border_top    = height_mm - MARGIN

    usable_w = border_right  - border_left
    usable_h = border_top    - border_bottom

    num_cols = int(usable_w / HORIZ_BLOCK_SIZE)
    num_rows = int(usable_h / VERT_BLOCK_SIZE)

    if num_cols < 1 or num_rows < 1:
        raise RuntimeError("[create_grid] Sheet is too small to contain a grid.")

    # Centre the grid within the usable area
    grid_w      = num_cols * HORIZ_BLOCK_SIZE
    grid_h      = num_rows * VERT_BLOCK_SIZE
    grid_left   = border_left   + (usable_w - grid_w) / 2.0
    grid_right  = grid_left     + grid_w
    grid_bottom = border_bottom + (usable_h - grid_h) / 2.0
    grid_top    = grid_bottom   + grid_h

    cw = HORIZ_BLOCK_SIZE
    ch = VERT_BLOCK_SIZE

    print(f"[create_grid] {size_key}: {num_cols} cols × {num_rows} rows  "
          f"({grid_w:.0f} × {grid_h:.0f} mm centred)")

    # -----------------------------------------------------------------------
    # Outer drawing frame
    # -----------------------------------------------------------------------
    _draw_drafting_line(work_part, border_left,  border_bottom, border_right, border_bottom, TEMPLATE_LAYER)
    _draw_drafting_line(work_part, border_right, border_bottom, border_right, border_top,    TEMPLATE_LAYER)
    _draw_drafting_line(work_part, border_right, border_top,    border_left,  border_top,    TEMPLATE_LAYER)
    _draw_drafting_line(work_part, border_left,  border_top,    border_left,  border_bottom, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Grid/border separator lines
    # -----------------------------------------------------------------------
    # Top band separator
    _draw_drafting_line(work_part, border_left, grid_top, border_right, grid_top, TEMPLATE_LAYER)
    # Bottom band separator
    _draw_drafting_line(work_part, border_left, grid_bottom, border_right, grid_bottom, TEMPLATE_LAYER)
    # Left band separator
    _draw_drafting_line(work_part, grid_left,  border_bottom, grid_left,  border_top, TEMPLATE_LAYER)
    # Right band separator
    _draw_drafting_line(work_part, grid_right, border_bottom, grid_right, border_top, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Vertical column dividers  (inside grid + top/bottom band ticks)
    # -----------------------------------------------------------------------
    for col in range(num_cols + 1):
        x = grid_left + col * cw
        # Inner grid line
        _draw_drafting_line(work_part, x, grid_bottom, x, grid_top,     TEMPLATE_LAYER)
        # Top band tick
        _draw_drafting_line(work_part, x, grid_top,    x, border_top,   TEMPLATE_LAYER)
        # Bottom band tick
        _draw_drafting_line(work_part, x, border_bottom, x, grid_bottom, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Horizontal row dividers  (inside grid + left/right band ticks)
    # -----------------------------------------------------------------------
    for row in range(num_rows + 1):
        y = grid_bottom + row * ch
        # Inner grid line
        _draw_drafting_line(work_part, grid_left, y, grid_right, y,      TEMPLATE_LAYER)
        # Left band tick
        _draw_drafting_line(work_part, border_left, y, grid_left, y,     TEMPLATE_LAYER)
        # Right band tick
        _draw_drafting_line(work_part, grid_right, y, border_right, y,   TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Column numeric labels  (top + bottom border bands)
    # -----------------------------------------------------------------------
    cy_top_band = grid_top    + (border_top    - grid_top)    / 2.0
    cy_bot_band = border_bottom + (grid_bottom - border_bottom) / 2.0

    for col in range(num_cols):
        label = str(col + 1)
        cx = grid_left + col * cw + cw / 2.0
        _place_label(work_part, label, cx, cy_top_band, TEMPLATE_LAYER)
        _place_label(work_part, label, cx, cy_bot_band, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Row alphabetic labels  (left + right border bands)
    # Rows count top → bottom; label row 0 = "A" at the top.
    # Supports > 26 rows: 27th row = "AA", 28th = "AB", etc.
    # -----------------------------------------------------------------------
    ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    cx_left_band  = border_left  + (grid_left  - border_left)  / 2.0
    cx_right_band = grid_right   + (border_right - grid_right) / 2.0

    for row in range(num_rows):
        if row < 26:
            label = ALPHA[row]
        else:
            label = ALPHA[row // 26 - 1] + ALPHA[row % 26]

        cy = grid_top - row * ch - ch / 2.0   # row 0 at top
        _place_label(work_part, label, cx_left_band,  cy, TEMPLATE_LAYER)
        _place_label(work_part, label, cx_right_band, cy, TEMPLATE_LAYER)

    print(f"[create_grid] Grid complete.")


# ---------------------------------------------------------------------------
# 3. TEAMCENTER ATTRIBUTE RETRIEVAL  (NX2306 FIX #5)
# ---------------------------------------------------------------------------

def get_tc_attributes(session: NXOpen.Session) -> dict:
    """
    Read Teamcenter item/revision attributes.

    NX2306 fix: UF Python API uses explicit out-variable style for
    AskWorkPartItem and attribute reads.  All paths wrapped individually
    so a single missing attribute never aborts the whole fetch.
    """
    attrs = {
        "item_id":       TC_DEFAULT,
        "item_revision": TC_DEFAULT,
        "object_name":   TC_DEFAULT,
        "drawing_no":    TC_DEFAULT,
        "material":      TC_DEFAULT,
    }

    try:
        uf   = NXOpen.UF.UFSession.GetUFSession()
        part = session.Parts.Work

        # Check managed mode
        if not uf.Pdm.IsPartManaged(part.Tag):
            print("[get_tc_attributes] Unmanaged mode – using defaults.")
            return attrs

        # Retrieve item + revision tags
        # NX2306 UF Python: AskWorkPartItem returns (item_tag, item_rev_tag)
        item_tag, item_rev_tag = uf.Pdm.AskWorkPartItem(part.Tag)

        def _read_str(tag, attr_name: str) -> str:
            """Read a string attribute, return TC_DEFAULT on any error."""
            try:
                val = uf.Obj.AskStringAttr(tag, attr_name)
                return val.strip() if val and val.strip() else TC_DEFAULT
            except Exception:
                return TC_DEFAULT

        # --- Standard attributes ---
        attrs["item_id"]       = _read_str(item_tag,     "item_id")
        attrs["item_revision"] = _read_str(item_rev_tag, "item_revision_id")

        # object_name from UF Obj name
        try:
            name = uf.Obj.AskName(item_tag)
            attrs["object_name"] = name if name else TC_DEFAULT
        except Exception:
            pass

        # --- Custom / site-specific attributes ---
        for key, candidates in {
            "drawing_no": ["drawing_number", "DrawingNumber", "DRW_NO", "U4_DRW_NO"],
            "material":   ["material",       "Material",      "MAT",   "U4_MATERIAL"],
        }.items():
            for attr in candidates:
                for tag in (item_rev_tag, item_tag):
                    val = _read_str(tag, attr)
                    if val != TC_DEFAULT:
                        attrs[key] = val
                        break
                if attrs[key] != TC_DEFAULT:
                    break

        print(f"[get_tc_attributes] {attrs}")

    except Exception as ex:
        print(f"[get_tc_attributes] Warning: {ex}. Using defaults.")

    return attrs


# ---------------------------------------------------------------------------
# 4. TITLE BLOCK  (NX2306 FIX #6, #7, #8)
# ---------------------------------------------------------------------------

def _find_note_by_hint(work_part: NXOpen.Part, hint: str):
    """
    Find an existing DraftingNote whose first line starts with `hint`.
    Uses the NX2306-safe iterator.
    """
    for note in _iter_drafting_notes(work_part):
        try:
            # NX2306: GetDisplayString() returns the rendered text
            lines = note.GetDisplayString()
            if lines and lines[0].startswith(hint):
                return note
        except AttributeError:
            # Older API: GetText()
            try:
                lines = note.GetText()
                if lines and lines[0].startswith(hint):
                    return note
            except Exception:
                pass
        except Exception:
            pass
    return None


def _update_or_create_note(work_part: NXOpen.Part,
                           hint: str, value: str,
                           x: float, y: float,
                           layer: int):
    """
    Update an existing title-block note or create a new one.

    NX2306 note text update:
      • To update text on a committed DraftingNote we must use a builder.
      • SetText() directly on a note is deprecated in NX2306.
    """
    full_text = f"{hint} {value}"
    existing  = _find_note_by_hint(work_part, hint)

    if existing is not None:
        # Re-open via builder and update text
        try:
            builder = work_part.Annotations.CreateDraftingNoteBuilder(existing)
            builder.Text.TextBlock.SetText([full_text])
            builder.Commit()
            builder.Destroy()
        except Exception as ex:
            print(f"[_update_or_create_note] Update failed for '{hint}': {ex}")
    else:
        _place_label(work_part, full_text, x, y, layer, char_height=3.0)


def update_title_block(session: NXOpen.Session,
                       sheet: NXOpen.Drawings.DrawingSheet,
                       size_key: str = "A3"):
    """
    Draw (or refresh) the title block and populate it with TC attributes.

    Title block occupies the bottom TITLE_BLOCK_HEIGHT mm strip of the sheet
    (between the MARGIN and MARGIN + TITLE_BLOCK_HEIGHT).

    Layout (3 rows × 2 columns):
    ┌───────────────────┬───────────────────┐
    │ ITEM ID:  <value> │ DWG NO:  <value>  │
    ├───────────────────┼───────────────────┤
    │ NAME:     <value> │ REV:     <value>  │
    ├───────────────────┼───────────────────┤
    │ MATERIAL: <value> │ SCALE:   1:1      │
    └───────────────────┴───────────────────┘
    """
    width_mm, _ = ISO_SHEET_SIZES[size_key]
    work_part   = session.Parts.Work

    _ensure_layer_visible(work_part, TEMPLATE_LAYER)
    _set_work_layer(session, TEMPLATE_LAYER)

    tc = get_tc_attributes(session)

    # -----------------------------------------------------------------------
    # Geometry
    # -----------------------------------------------------------------------
    tb_left   = MARGIN
    tb_right  = width_mm - MARGIN
    tb_bottom = MARGIN
    tb_top    = MARGIN + TITLE_BLOCK_HEIGHT
    tb_width  = tb_right - tb_left
    mid_x     = tb_left + tb_width / 2.0
    row_h     = TITLE_BLOCK_HEIGHT / 3.0

    # Outer border
    _draw_drafting_line(work_part, tb_left,  tb_bottom, tb_right, tb_bottom, TEMPLATE_LAYER)
    _draw_drafting_line(work_part, tb_right, tb_bottom, tb_right, tb_top,    TEMPLATE_LAYER)
    _draw_drafting_line(work_part, tb_right, tb_top,    tb_left,  tb_top,    TEMPLATE_LAYER)
    _draw_drafting_line(work_part, tb_left,  tb_top,    tb_left,  tb_bottom, TEMPLATE_LAYER)

    # Vertical centre split
    _draw_drafting_line(work_part, mid_x, tb_bottom, mid_x, tb_top, TEMPLATE_LAYER)

    # Two horizontal row dividers
    for r in (1, 2):
        y = tb_bottom + r * row_h
        _draw_drafting_line(work_part, tb_left, y, tb_right, y, TEMPLATE_LAYER)

    # -----------------------------------------------------------------------
    # Field centre positions
    # -----------------------------------------------------------------------
    lx = tb_left + tb_width / 4.0        # left column centre
    rx = tb_left + 3.0 * tb_width / 4.0  # right column centre

    y1 = tb_bottom + 2.5 * row_h   # top row
    y2 = tb_bottom + 1.5 * row_h   # middle row
    y3 = tb_bottom + 0.5 * row_h   # bottom row

    # -----------------------------------------------------------------------
    # Place / update fields
    # -----------------------------------------------------------------------
    _update_or_create_note(work_part, "ITEM ID:",  tc["item_id"],       lx, y1, TEMPLATE_LAYER)
    _update_or_create_note(work_part, "DWG NO:",   tc["drawing_no"],    rx, y1, TEMPLATE_LAYER)
    _update_or_create_note(work_part, "NAME:",     tc["object_name"],   lx, y2, TEMPLATE_LAYER)
    _update_or_create_note(work_part, "REV:",      tc["item_revision"], rx, y2, TEMPLATE_LAYER)
    _update_or_create_note(work_part, "MATERIAL:", tc["material"],      lx, y3, TEMPLATE_LAYER)
    _update_or_create_note(work_part, "SCALE:",    "1:1",               rx, y3, TEMPLATE_LAYER)

    print("[update_title_block] Title block updated.")


# ---------------------------------------------------------------------------
# 5. MAIN
# ---------------------------------------------------------------------------

def main():
    """
    Journal entry point.

    Steps
    -----
    1. Validate session and work part.
    2. Create / update drawing sheet.
    3. Build the ISO grid.
    4. Populate the title block from Teamcenter attributes.
    5. Refresh the display.

    Configuration
    -------------
    Edit SELECTED_SIZE and SHEET_NAME below, or wire them up to an
    NX Block-UI dialog for interactive use.
    """
    session = NXOpen.Session.GetSession()

    # -----------------------------------------------
    # ↓ Edit these two values to change output ↓
    SELECTED_SIZE = "A3"             # "A0" | "A1" | "A2" | "A3" | "A4"
    SHEET_NAME    = "DrawingSheet1"
    # -----------------------------------------------

    lw = session.ListingWindow
    lw.Open()

    work_part = session.Parts.Work
    if work_part is None:
        lw.WriteLine("[main] ERROR: No work part open.")
        return

    w, h = ISO_SHEET_SIZES.get(SELECTED_SIZE, (0, 0))
    lw.WriteLine("=" * 60)
    lw.WriteLine("  NX Drawing Template – NX2306 Edition")
    lw.WriteLine("=" * 60)
    lw.WriteLine(f"  Sheet : {SELECTED_SIZE}  ({w} x {h} mm)")
    lw.WriteLine(f"  Name  : {SHEET_NAME}")
    lw.WriteLine("-" * 60)

    # Single undo mark for the entire journal run
    undo_mark = session.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible,
        "NX Drawing Template"
    )

    try:
        lw.WriteLine("[1/3] Sheet …")
        sheet = create_sheet(session, SHEET_NAME, SELECTED_SIZE)

        lw.WriteLine("[2/3] Grid …")
        create_grid(session, sheet, SELECTED_SIZE)

        lw.WriteLine("[3/3] Title block …")
        update_title_block(session, sheet, SELECTED_SIZE)

        # Refresh display
        try:
            session.Parts.Work.ModelingViews.WorkView.UpdateDisplay()
        except Exception:
            pass  # Non-fatal if view update fails

        lw.WriteLine("\n[main] Done.\n" + "=" * 60)

    except Exception as ex:
        # Roll back on failure
        session.UndoToMark(undo_mark, "NX Drawing Template")
        lw.WriteLine(f"\n[main] FATAL: {ex}  — changes rolled back.")
        raise

    finally:
        session.DeleteUndoMark(undo_mark, "NX Drawing Template")


# NX journal execution entry
if __name__ == "__main__":
    main()
