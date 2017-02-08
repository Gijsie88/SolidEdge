Imports System.Runtime.InteropServices

Module Module1

    Sub Main()
        Dim oApp As SolidEdgeFramework.Application
        Dim oDoc As SolidEdgeDraft.DraftDocument
        Dim oSel As SolidEdgeFramework.SelectSet
        Dim iCount As Integer

        oApp = Marshal.GetActiveObject("SolidEdge.Application")
        oDoc = oApp.ActiveDocument
        oSel = oDoc.SelectSet
        iCount = oSel.Count.ToString()

        oApp.StatusBar = iCount.ToString() + " object(s) selected."

    End Sub

End Module
