Imports Light_Bar

Public Class Form2
    Private WithEvents _client As New Light_Bar.MenuClient()

    Private Sub FormApp1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _client.StartListening()

        Dim meinMenue As New List(Of MenuItemData)

        ' - Datei
        Dim mDatei = New MenuItemData("Datei", "id_datei")

        ' Ebene 2: Speichern
        Dim mSpeichern = New MenuItemData("Speichern", "id_speichern")

        ' Ebene 3: Als Text *neuer container
        Dim mAlsText = New MenuItemData("Als Text", "id_als_text")

        ' Ebene 4: PDF unterpunkt
        mAlsText.SubItems.Add(New MenuItemData("PDF", "cmd_save_pdf"))

        ' Als Text Speichern 
        mSpeichern.SubItems.Add(mAlsText)

        ' Speichern zu Datei
        mDatei.SubItems.Add(mSpeichern)

        ' - Bearbeiten
        Dim mBearbeiten = New MenuItemData("Bearbeiten", "id_bearbeiten")
        Dim mText = New MenuItemData("Text", "id_text")
        mText.SubItems.Add(New MenuItemData("Kopieren", "cmd_copy_text"))
        mBearbeiten.SubItems.Add(mText)

        ' Weiterleitung
        meinMenue.Add(mDatei)
        meinMenue.Add(mBearbeiten)

        _client.SendMenu(meinMenue)
    End Sub

    Private Sub _client_MenuClicked(commandID As String) Handles _client.MenuClicked
        If Me.InvokeRequired Then
            Me.Invoke(Sub() _client_MenuClicked(commandID))
            Return
        End If

        Select Case commandID
            Case "cmd_save_txt"
                SaveAs()

            Case "cmd_save_pdf"
                SaveAsPDF()

            Case "cmd_copy_text"
                TextCopy()

            Case Else
                Debug.WriteLine("Unknown Command: " & commandID)
        End Select
    End Sub

    Private Sub SaveAs()
        MessageBox.Show("Success...")
    End Sub

    Private Sub SaveAsPDF()
        MessageBox.Show("Success...")
    End Sub

    Private Sub TextCopy()
        MessageBox.Show("Success...")
    End Sub
End Class