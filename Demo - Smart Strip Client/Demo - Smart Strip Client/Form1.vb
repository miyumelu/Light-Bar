Imports System.IO
Imports System.IO.Pipes
Imports System.Text.Json
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class Form1
    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    Private appMenus As New Dictionary(Of Integer, List(Of MenuItemData2))
    Private lastProcessID As Integer = -1

    Private Sub FormDesktop_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Bounds = New Rectangle(0, 0, Screen.PrimaryScreen.Bounds.Width, 30)

        Task.Run(Sub() ListenForApps())

        TimerFocus.Interval = 250
        TimerFocus.Start()

        ShowDefaultMenu()
    End Sub

    Private Sub ShowDefaultMenu()
        MenuStrip1.Items.Clear()
        Dim item As New ToolStripMenuItem("Finder")
        item.DropDownItems.Add("Über diesen PC")
        MenuStrip1.Items.Add(item)
    End Sub

    Private Sub TimerFocus_Tick(sender As Object, e As EventArgs) Handles TimerFocus.Tick
        Dim hwnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        GetWindowThreadProcessId(hwnd, procID)

        If procID = Process.GetCurrentProcess().Id Then
            Exit Sub
        End If

        If appMenus.ContainsKey(procID) Then
            If procID <> lastProcessID Then
                lastProcessID = procID
                BuildFinalMenu(appMenus(procID))
            End If
        Else
            If lastProcessID <> -1 Then
                If Not IsProcessRunning(lastProcessID) Then
                    appMenus.Remove(lastProcessID)
                    lastProcessID = -1
                    ShowDefaultMenu()
                Else
                    lastProcessID = -1
                    ShowDefaultMenu()
                End If
            End If
        End If
    End Sub

    Private Function IsProcessRunning(pid As Integer) As Boolean
        Try
            Dim p As Process = Process.GetProcessById(pid)
            Return Not p.HasExited
        Catch ex As ArgumentException
            Return False
        End Try
    End Function
    Private Sub ListenForApps()
        While True
            Try
                Using server As New NamedPipeServerStream("MacOSMenuPipe", PipeDirection.In)
                    server.WaitForConnection()
                    Using reader As New StreamReader(server)
                        Dim json = reader.ReadToEnd()
                        Dim package = JsonSerializer.Deserialize(Of MenuPackage2)(json)
                        appMenus(package.ProcessID) = package.MenuItems
                    End Using
                End Using
            Catch : End Try
        End While
    End Sub

    Private Sub BuildFinalMenu(data As List(Of MenuItemData2))
        If Me.InvokeRequired Then
            Me.Invoke(Sub() BuildFinalMenu(data))
            Return
        End If
        MenuStrip1.Items.Clear()
        For Each item In data
            MenuStrip1.Items.Add(CreateMenuItemRecursive(item))
        Next
    End Sub

    Private Function CreateMenuItemRecursive(data As MenuItemData2) As ToolStripMenuItem
        Dim tsItem As New ToolStripMenuItem(data.Text)
        tsItem.Tag = data.ID
        For Each subItem In data.SubItems
            tsItem.DropDownItems.Add(CreateMenuItemRecursive(subItem))
        Next
        AddHandler tsItem.Click, AddressOf RemoteMenu_Click
        Return tsItem
    End Function

    Private Sub RemoteMenu_Click(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)
        If item.HasDropDownItems Then Return

        Dim cmdID = item.Tag.ToString()
        Task.Run(Sub()
                     Try
                         Using client As New NamedPipeClientStream(".", "MacOSClickPipe_" & lastProcessID, PipeDirection.Out)
                             client.Connect(200)
                             Using writer As New StreamWriter(client)
                                 writer.Write(cmdID)
                             End Using
                         End Using
                     Catch : End Try
                 End Sub)
    End Sub
End Class