Imports System.Drawing
Imports System.Threading

Public Class frmMain
#Region "DECLARATIONS"
    Private SQL As New SQLControl
    Public Version As String = "1.5.0"
    Public PowFactor As Single = 1
    Public Index As Integer

    Private MyGenset As Genset
    Private GensetList As New List(Of Genset)

    Public bmp As Bitmap
#End Region

#Region "QUERY SYNTHESIS"
    Private Sub SynthQuery()

    End Sub
#End Region

#Region "FORM OBJECTS"
#Region "-> Filter Tab"
    ' ALL FILTER TAB CODE HERE
#End Region
#Region "-> Heat Tab"
    ' ALL HEAT RECOVERY CODE HERE
#End Region
#Region "-> Compare Tab"
    ' ALL COMPARE CODE HERE
#End Region
#Region "-> View Tab"
    ' ALL VIEW CODE HERE
#End Region
#Region "-> Final Tab"
    ' ALL PRINT/SAVE CODE HERE
#End Region
#Region "-> Buttons"
    ' ALL BUTTON CODE HERE
#End Region
#Region "-> Menu Strip"
    Private Sub miDB_Click(sender As System.Object, e As System.EventArgs) Handles miDB.Click
        Dim EditDB As New frmEditor : EditDB.Show()
    End Sub
#End Region
#End Region
End Class