Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
 
Module Module1
 
    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
 
    Dim theUI As UI = UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow
 
    Sub Main()
 
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "NXJ")
 
        lw.Open()
   
        Dim notesOnSheet As New List(Of Annotations.Note)
    Dim xc,xd As Integer
    xc = 0
    xd = 0
        'lw.WriteLine("working with drawing sheet: " & theSession.Parts.Work.DrawingSheets.CurrentDrawingSheet.Name)
 
        For Each temp As DisplayableObject In theSession.Parts.Work.DrawingSheets.CurrentDrawingSheet.View.AskVisibleObjects
            If TypeOf (temp) Is Annotations.Note Then
                Dim theNote As Annotations.Note = temp
                For i As Integer = 0 To theNote.GetText.Length - 1
                    If theNote.GetText(i).Contains("CC") Then
                        xc=xc+1 
                      
                    End If           
                       
                Next
                For j As Integer = 0 To theNote.GetText.Length - 1
                    If theNote.GetText(j).Contains("SC") Then
                        xd=xd+1 
                      
                    End If           
                       
                Next
               
            End If
        Next
        lw.WriteLine("Count  CC == "& xc &"  & SC ==" & xd)
        lw.Close()
 
    End Sub
 
 
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
 
    End Function
 
End Module

