Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF

Module Symbolqty

    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()

    Dim theUI As UI = UI.GetUI()

    Sub Main()

        ' Set an undo mark for error handling
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Swapnil")

        ' Create lists to store notes on the sheet and the counts of SC and CC symbols
        Dim notesOnSheet As New List(Of Annotations.Note)
        Dim oSc, oCc As Integer
        oSc = 0
        oCc = 0

        ' Iterate through visible objects on the current drawing sheet
        For Each temp As DisplayableObject In theSession.Parts.Work.DrawingSheets.CurrentDrawingSheet.View.AskVisibleObjects
            If TypeOf (temp) Is Annotations.Note Then

                ' Check if the object is a Note
                Dim theNote As Annotations.Note = temp

                ' Iterate through each line of text in the note
                For i As Integer = 0 To theNote.GetText.Length - 1
                    ' Check if the line contains "CC"
                    If theNote.GetText(i).Contains("CC") Then
                        oCc = oCc + 1
                    End If
                Next

                ' Iterate through each line of text in the note
                For j As Integer = 0 To theNote.GetText.Length - 1
                    ' Check if the line contains "SC"
                    If theNote.GetText(j).Contains("SC") Then
                        oSc = oSc + 1
                    End If
                Next

            End If
        Next

        ' Display the counts of SC and CC symbols
        MsgBox("SC = " & oSc & "  And  CC = " & oCc, Title:="Symbol Qty")

    End Sub


    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        ' Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

    End Function

End Module
