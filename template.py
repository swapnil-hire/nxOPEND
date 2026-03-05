# NXOpen Python Journal
# Sheet + Grid + Teamcenter Title Block Automation

import NXOpen
import NXOpen.UF
import string
import math

# --------------------------------------------
# ISO SHEET SIZE DICTIONARY (mm)
# --------------------------------------------
ISO_SIZES = {
    "A0": (1189.0, 841.0),
    "A1": (841.0, 594.0),
    "A2": (594.0, 420.0),
    "A3": (420.0, 297.0),
    "A4": (297.0, 210.0),
}

MARGIN = 10.0
H_BLOCK = 30.0
V_BLOCK = 24.0


# --------------------------------------------
# MAIN FUNCTION
# --------------------------------------------
def main():

    session = NXOpen.Session.GetSession()
    workPart = session.Parts.Work
    ufSession = NXOpen.UF.UFSession.GetUFSession()

    try:

        sheet_size = "A3"  # <<< Change or connect to UI selection

        width, height = ISO_SIZES[sheet_size]

        create_sheet(workPart, width, height)
        create_grid(workPart, width, height)
        update_title_block(workPart)

        session.ListingWindow.Open()
        session.ListingWindow.WriteLine("Template customization completed successfully.")

    except Exception as e:
        session.ListingWindow.Open()
        session.ListingWindow.WriteLine("Error: " + str(e))


# --------------------------------------------
# CREATE OR UPDATE DRAWING SHEET
# --------------------------------------------
def create_sheet(workPart, width, height):

    sheets = workPart.DrawingSheets

    if sheets.Count > 0:
        sheet = sheets.ToArray()[0]
        sheet.SetSize(NXOpen.DrawingSheet.Unit.Millimeter, width, height)
    else:
        sheetBuilder = workPart.DrawingSheets.DrawingSheetBuilder(NXOpen.DrawingSheet.Null)
        sheetBuilder.Height = height
        sheetBuilder.Length = width
        sheetBuilder.Commit()
        sheetBuilder.Destroy()


# --------------------------------------------
# GRID CREATION FUNCTION
# --------------------------------------------
def create_grid(workPart, width, height):

    grid_width = width - 2 * MARGIN
    grid_height = height - 2 * MARGIN

    num_horizontal = int(grid_width // H_BLOCK)
    num_vertical = int(grid_height // V_BLOCK)

    curves = workPart.Curves
    annotations = workPart.Annotations

    # Vertical grid lines
    for i in range(num_horizontal + 1):
        x = MARGIN + i * H_BLOCK
        start = NXOpen.Point3d(x, MARGIN, 0)
        end = NXOpen.Point3d(x, MARGIN + num_vertical * V_BLOCK, 0)
        curves.CreateLine(start, end)

    # Horizontal grid lines
    for j in range(num_vertical + 1):
        y = MARGIN + j * V_BLOCK
        start = NXOpen.Point3d(MARGIN, y, 0)
        end = NXOpen.Point3d(MARGIN + num_horizontal * H_BLOCK, y, 0)
        curves.CreateLine(start, end)

    # Horizontal numbering (1,2,3...)
    for i in range(num_horizontal):
        x_center = MARGIN + (i + 0.5) * H_BLOCK
        y_pos = MARGIN - 5
        noteText = str(i + 1)
        create_note(workPart, noteText, x_center, y_pos)

    # Vertical lettering (A,B,C...)
    letters = list(string.ascii_uppercase)

    for j in range(min(num_vertical, 26)):
        y_center = MARGIN + (j + 0.5) * V_BLOCK
        x_pos = MARGIN - 8
        create_note(workPart, letters[j], x_pos, y_center)


# --------------------------------------------
# CREATE NOTE HELPER
# --------------------------------------------
def create_note(workPart, text, x, y):

    noteBuilder = workPart.Annotations.CreateNoteBuilder(NXOpen.Annotations.Note.Null)

    noteBuilder.Text.TextBlock.SetText(text)
    noteBuilder.Origin = NXOpen.Point3d(x, y, 0)

    note = noteBuilder.Commit()
    noteBuilder.Destroy()


# --------------------------------------------
# FETCH TEAMCENTER ATTRIBUTES
# --------------------------------------------
def get_tc_attributes(workPart):

    attributes = {}

    try:
        attributes["ItemID"] = workPart.GetStringAttribute("DB_PART_NO")
    except:
        attributes["ItemID"] = "-"

    try:
        attributes["Revision"] = workPart.GetStringAttribute("DB_PART_REV")
    except:
        attributes["Revision"] = "-"

    try:
        attributes["ObjectName"] = workPart.GetStringAttribute("DB_PART_NAME")
    except:
        attributes["ObjectName"] = "-"

    try:
        attributes["DrawingNumber"] = workPart.GetStringAttribute("DB_DRAWING_NO")
    except:
        attributes["DrawingNumber"] = "-"

    try:
        attributes["Material"] = workPart.GetStringAttribute("DB_MATERIAL")
    except:
        attributes["Material"] = "-"

    return attributes


# --------------------------------------------
# UPDATE TITLE BLOCK NOTES
# --------------------------------------------
def update_title_block(workPart):

    attrs = get_tc_attributes(workPart)

    # Example fixed title block positions (modify as per your template)
    base_x = 50
    base_y = 20

    create_note(workPart, "Item ID: " + attrs["ItemID"], base_x, base_y)
    create_note(workPart, "Revision: " + attrs["Revision"], base_x, base_y + 8)
    create_note(workPart, "Name: " + attrs["ObjectName"], base_x, base_y + 16)
    create_note(workPart, "Drawing No: " + attrs["DrawingNumber"], base_x, base_y + 24)
    create_note(workPart, "Material: " + attrs["Material"], base_x, base_y + 32)


# --------------------------------------------
# NX ENTRY POINT
# --------------------------------------------
if __name__ == '__main__':
    main()
