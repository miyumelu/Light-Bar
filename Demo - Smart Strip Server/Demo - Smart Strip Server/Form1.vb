Imports System.IO
Imports System.IO.Pipes
Imports System.Text.Json
Imports System.Threading

Public Class Form1
    Private myID As Integer = Process.GetCurrentProcess().Id

    Private Sub FormApp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SendMenuToDesktop()
        Task.Run(Sub() ListenForClicks())
    End Sub

    Private Sub SendMenuToDesktop()
        Dim package As New MenuPackage2 With {.ProcessID = myID}

        Dim fileMenu As New MenuItemData2("Datei", "m_file")
        fileMenu.SubItems.Add(New MenuItemData2("Neu", "cmd_new"))

        Dim exportSub As New MenuItemData2("Exportieren", "m_exp")
        exportSub.SubItems.Add(New MenuItemData2("als PDF", "cmd_pdf"))
        exportSub.SubItems.Add(New MenuItemData2("als JPG", "cmd_jpg"))

        fileMenu.SubItems.Add(exportSub)
        package.MenuItems.Add(fileMenu)
        package.MenuItems.Add(New MenuItemData2("Hilfe", "cmd_help"))

        Try
            Using client As New NamedPipeClientStream(".", "MacOSMenuPipe", PipeDirection.Out)
                client.Connect(1000)
                Dim json = JsonSerializer.Serialize(package)
                Using writer As New StreamWriter(client)
                    writer.Write(json)
                End Using
            End Using
        Catch
            Me.Text = "Strip not found"
        End Try
    End Sub

    Private Sub ListenForClicks()
        While True
            Try
                Using server As New NamedPipeServerStream("MacOSClickPipe_" & myID, PipeDirection.In)
                    server.WaitForConnection()
                    Using reader As New StreamReader(server)
                        Dim cmdID = reader.ReadToEnd()
                        Me.Invoke(Sub() HandleAction(cmdID))
                    End Using
                End Using
            Catch : End Try
        End While
    End Sub

    Private Sub HandleAction(id As String)
        Select Case id
            Case "cmd_pdf" : MessageBox.Show("Success...")
            Case "cmd_help" : MessageBox.Show("Success...")
            Case "cmd_new" : Me.BackColor = Color.Red
        End Select
    End Sub
End Class