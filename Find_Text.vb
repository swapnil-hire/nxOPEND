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
 
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Swapnil")
 
       
   
        Dim notesOnSheet As New List(Of Annotations.Note)
        Dim oSc,oCc As Integer
        oSc = 0
        oCc = 0
        
 
        For Each temp As DisplayableObject In theSession.Parts.Work.DrawingSheets.CurrentDrawingSheet.View.AskVisibleObjects
            If TypeOf (temp) Is Annotations.Note Then

                Dim theNote As Annotations.Note = temp

                For i As Integer = 0 To theNote.GetText.Length - 1
                    If theNote.GetText(i).Contains("CC") Then
                        oCc=oCc+1 
                      
                    End If           
                       
                Next
                For j As Integer = 0 To theNote.GetText.Length - 1
                    If theNote.GetText(j).Contains("SC") Then
                       oSc=oSc+1 
                      
                    End If           
                       
                Next
               
            End If
        Next
        MsgBox("SC = "& oSc &"  And  CC = " & oCc , Title:="Symbol Qty" )

         
    End Sub
 
 
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
 
    End Function

End Module


