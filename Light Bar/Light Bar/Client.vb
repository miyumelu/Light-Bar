Imports System.IO
Imports System.IO.Pipes
Imports System.Text.Json
Imports System.Threading

Public Class MenuClient
    Private _myID As Integer = Process.GetCurrentProcess().Id
    Public Event MenuClicked(commandID As String)

    Public Sub SendMenu(items As List(Of MenuItemData))
        Dim package As New MenuPackage With {.ProcessID = _myID, .MenuItems = items}
        Task.Run(Sub()
                     Try
                         Using client As New NamedPipeClientStream(".", "MacOSMenuPipe", PipeDirection.Out)
                             client.Connect(1000)
                             Dim json = JsonSerializer.Serialize(package)
                             Using writer As New StreamWriter(client)
                                 writer.Write(json)
                             End Using
                         End Using
                     Catch : End Try
                 End Sub)
    End Sub

    Public Sub StartListening()
        Dim t As New Thread(AddressOf ListenLoop)
        t.IsBackground = True
        t.Start()
    End Sub

    Private Sub ListenLoop()
        While True
            Try
                Using server As New NamedPipeServerStream("MacOSClickPipe_" & _myID, PipeDirection.In)
                    server.WaitForConnection()
                    Using reader As New StreamReader(server)
                        Dim cmdID = reader.ReadToEnd()
                        RaiseEvent MenuClicked(cmdID)
                    End Using
                End Using
            Catch : End Try
        End While
    End Sub
End Class