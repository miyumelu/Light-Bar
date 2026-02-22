Imports System.Collections.Generic

Public Class MenuPackage2
    Public Property ProcessID As Integer
    Public Property MenuItems As New List(Of MenuItemData2)
End Class

Public Class MenuItemData2
    Public Property Text As String
    Public Property ID As String
    Public Property SubItems As New List(Of MenuItemData2)

    Public Sub New()
    End Sub

    Public Sub New(txt As String, identifier As String)
        Text = txt
        ID = identifier
    End Sub
End Class