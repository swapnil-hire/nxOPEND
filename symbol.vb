Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF

Module Symbolqty

    Dim theSession   As Session   = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI        As UI        = UI.GetUI()

    Sub Main()

        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Swapnil")

        ' ---------------------------------------------------------------
        ' BUG FIX 1: "Dim oSc, oCc As Integer" in VB declares oSc as
        ' Object and only oCc as Integer. Must declare separately.
        ' ---------------------------------------------------------------
        Dim oSc As Integer = 0
        Dim oCc As Integer = 0
        Dim notesFound As Integer = 0

        ' ---------------------------------------------------------------
        ' BUG FIX 2: Check if a drawing sheet is active before proceeding
        ' ---------------------------------------------------------------
        Dim currentSheet As Drawings.DrawingSheet = 
                theSession.Parts.Work.DrawingSheets.CurrentDrawingSheet

        If currentSheet Is Nothing Then
            MsgBox("No active drawing sheet found. Please open a drawing sheet first.",
                   MsgBoxStyle.Exclamation, "Symbol Qty")
            Return
        End If

        ' ---------------------------------------------------------------
        ' BUG FIX 3: Wrapped in Try/Catch to handle NX API errors safely
        ' ---------------------------------------------------------------
        Try

            For Each temp As DisplayableObject In currentSheet.View.AskVisibleObjects()

                If TypeOf temp Is Annotations.Note Then

                    Dim theNote As Annotations.Note = CType(temp, Annotations.Note)
                    Dim textLines() As String = theNote.GetText()

                    notesFound += 1

                    ' ---------------------------------------------------
                    ' FIX 4: Merged two separate loops into ONE loop
                    ' FIX 5: Count actual occurrences per line (not just
                    '         whether line contains the text). Handles
                    '         cases like "CC CC CC" = 3 counts.
                    ' FIX 6: Case-insensitive comparison added
                    ' ---------------------------------------------------
                    For Each line As String In textLines

                        Dim upperLine As String = line.ToUpper().Trim()

                        ' Count all occurrences of "SC" in this line
                        oSc += CountOccurrences(upperLine, "SC")

                        ' Count all occurrences of "CC" in this line
                        oCc += CountOccurrences(upperLine, "CC")

                    Next

                End If

            Next

        Catch ex As Exception
            MsgBox("Error during processing: " & ex.Message,
                   MsgBoxStyle.Critical, "Symbol Qty Error")
            Return
        End Try

        ' ---------------------------------------------------------------
        ' FIX 7: Improved result message with note count and sheet name
        ' ---------------------------------------------------------------
        Dim resultMsg As String =
            "Drawing Sheet : " & currentSheet.Name                  & vbNewLine &
            "Total Notes   : " & notesFound.ToString()               & vbNewLine &
            "─────────────────────────────"                          & vbNewLine &
            "SC Count       : " & oSc.ToString()                     & vbNewLine &
            "CC Count       : " & oCc.ToString()                     & vbNewLine &
            "─────────────────────────────"                          & vbNewLine &
            "Total Symbols  : " & (oSc + oCc).ToString()

        MsgBox(resultMsg, MsgBoxStyle.Information, "Symbol Qty")

    End Sub

    ' =======================================================================
    ' HELPER: Count how many times a keyword appears in a line of text
    ' Handles multiple occurrences in the same line e.g. "CC and CC" = 2
    ' =======================================================================
    Function CountOccurrences(sourceLine As String, keyword As String) As Integer

        If String.IsNullOrEmpty(sourceLine) OrElse
           String.IsNullOrEmpty(keyword) Then Return 0

        Dim count    As Integer = 0
        Dim index    As Integer = 0
        Dim position As Integer

        Do
            position = sourceLine.IndexOf(keyword, index, StringComparison.OrdinalIgnoreCase)
            If position = -1 Then Exit Do
            count  += 1
            index   = position + keyword.Length
        Loop

        Return count

    End Function

    ' =======================================================================
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

End Module
