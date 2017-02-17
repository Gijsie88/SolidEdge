Imports System.Runtime.InteropServices


Module Module1

    Sub Main()
        Dim oApp As SolidEdgeFramework.Application
        Dim oDoc As SolidEdgeAssembly.AssemblyDocument
        Dim oSel As SolidEdgeFramework.SelectSet
        Dim iCount As Integer

        oApp = Marshal.GetActiveObject("SolidEdge.Application")
        oDoc = oApp.ActiveDocument
        SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsReports


        MsgBox(oDoc.Name)

    End Sub

End Module
