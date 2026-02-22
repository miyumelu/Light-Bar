Imports System.IO
Imports System.IO.Pipes
Imports System.Text.Json
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class MenuServer
    <DllImport("user32.dll")> Private Shared Function GetForegroundWindow() As IntPtr : End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer : End Function

    Public Event RequestMenuUpdate(items As List(Of MenuItemData))
    Public Event RequestDefaultMenu()

    Private _appMenus As New Dictionary(Of Integer, List(Of MenuItemData))
    Private _lastProcessID As Integer = -1
    Private _myPID As Integer = Process.GetCurrentProcess().Id

    Public Sub Start()
        Dim t As New Thread(AddressOf ListenForApps)
        t.IsBackground = True
        t.Start()
    End Sub

    Public Sub CheckFocus()
        Dim hwnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        GetWindowThreadProcessId(hwnd, procID)

        If procID = _myPID Then
            If _lastProcessID <> -1 Then
                If Not IsProcessRunning(_lastProcessID) Then
                    _appMenus.Remove(_lastProcessID)
                    _lastProcessID = -1
                    RaiseEvent RequestDefaultMenu()
                End If
            End If
            Exit Sub
        End If

        If procID <> _lastProcessID Then
            _lastProcessID = procID
            If _appMenus.ContainsKey(procID) Then
                If IsProcessRunning(procID) Then
                    RaiseEvent RequestMenuUpdate(_appMenus(procID))
                Else
                    _appMenus.Remove(procID)
                    _lastProcessID = -1
                    RaiseEvent RequestDefaultMenu()
                End If
            Else
                RaiseEvent RequestDefaultMenu()
            End If
        End If
    End Sub

    Private Sub ListenForApps()
        While True
            Try
                Using server As New NamedPipeServerStream("MacOSMenuPipe", PipeDirection.In)
                    server.WaitForConnection()
                    Using reader As New StreamReader(server)
                        Dim json = reader.ReadToEnd()
                        Dim package = JsonSerializer.Deserialize(Of MenuPackage)(json)
                        _appMenus(package.ProcessID) = package.MenuItems
                        WatchProcess(package.ProcessID)
                        If package.ProcessID = _lastProcessID Then
                            RaiseEvent RequestMenuUpdate(package.MenuItems)
                        End If
                    End Using
                End Using
            Catch : End Try
        End While
    End Sub

    Private Sub WatchProcess(pid As Integer)
        Try
            Dim proc = Process.GetProcessById(pid)
            proc.EnableRaisingEvents = True
            AddHandler proc.Exited, Sub()
                                        _appMenus.Remove(pid)
                                        If _lastProcessID = pid Then
                                            _lastProcessID = -1
                                            RaiseEvent RequestDefaultMenu()
                                        End If
                                    End Sub
        Catch : End Try
    End Sub

    Public Sub SendClickToActiveApp(commandID As String)
        If _lastProcessID = -1 Then Exit Sub

        Dim targetPID = _lastProcessID
        Task.Run(Sub()
                     Try
                         Using client As New NamedPipeClientStream(".", "MacOSClickPipe_" & targetPID, PipeDirection.Out)
                             client.Connect(100)
                             Using writer As New StreamWriter(client)
                                 writer.Write(commandID)
                                 writer.Flush()
                             End Using
                         End Using
                     Catch ex As Exception
                         _appMenus.Remove(targetPID)
                         RaiseEvent RequestDefaultMenu()
                     End Try
                 End Sub)
    End Sub

    Private Function IsProcessRunning(pid As Integer) As Boolean
        Try
            Return Not Process.GetProcessById(pid).HasExited
        Catch : Return False : End Try
    End Function
End Class