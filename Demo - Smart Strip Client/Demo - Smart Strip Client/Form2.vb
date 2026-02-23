Imports System.Runtime.InteropServices
Imports Light_Bar

Public Class Form2
    Private WithEvents _server As New MenuServer()
    Private lastProcessID As Integer = -1

    Private adamBoldFont As New Font("Adam Bold", 16, FontStyle.Regular) 'Missing DCL connection
    Private adamMediumFont As New Font("Adam Medium", 16, FontStyle.Regular) 'Missing DCL connection

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim p As CreateParams = MyBase.CreateParams
            p.ExStyle = p.ExStyle Or &H8000000
            Return p
        End Get
    End Property

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        Me.FormBorderStyle = FormBorderStyle.None
        Me.TopMost = True
        Me.BackColor = Color.White
        Me.Bounds = New Rectangle(0, 0, Screen.PrimaryScreen.Bounds.Width, 30)

        Try
            _server.Start()
            Timer1.Interval = 200
            Timer1.Start()
            ShowDefaultMenu()
        Catch ex As Exception
            MessageBox.Show("Fehler beim Starten des Servers: " & ex.Message)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim hwnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        GetWindowThreadProcessId(hwnd, procID)

        Dim isMouseInSafeZone As Boolean = False
        Dim mousePos = Cursor.Position

        If MenuStrip1.Bounds.Contains(MenuStrip1.Parent.PointToClient(mousePos)) Then
            isMouseInSafeZone = True
        End If

        If Not isMouseInSafeZone Then
            isMouseInSafeZone = IsMouseOverAnyDropDown(MenuStrip1.Items)
        End If

        If procID <> Process.GetCurrentProcess().Id Then
            If Not isMouseInSafeZone Then
                CloseAllMenus()
                If procID <> lastProcessID AndAlso procID <> 0 Then
                    lastProcessID = procID
                    _server.CheckFocus()
                End If
            End If
        End If
    End Sub

    Private Function IsMouseOverAnyDropDown(items As ToolStripItemCollection) As Boolean
        Dim mousePos = Cursor.Position
        For Each item As ToolStripItem In items
            If TypeOf item Is ToolStripMenuItem Then
                Dim mi = DirectCast(item, ToolStripMenuItem)
                If mi.DropDown.Visible Then
                    If mi.DropDown.Bounds.Contains(mousePos) Then Return True
                    If IsMouseOverAnyDropDown(mi.DropDownItems) Then Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Sub _server_RequestMenuUpdate(items As List(Of MenuItemData)) Handles _server.RequestMenuUpdate
        If Me.InvokeRequired Then
            Me.Invoke(Sub() _server_RequestMenuUpdate(items))
            Return
        End If

        MenuStrip1.Items.Clear()

        For i As Integer = 0 To items.Count - 1
            Dim isFirstAppButton As Boolean = (i = 0)
            MenuStrip1.Items.Add(BuildMenuRecursive(items(i), True, isFirstAppButton))
        Next
    End Sub

    Private Sub _server_RequestDefaultMenu() Handles _server.RequestDefaultMenu
        If Me.InvokeRequired Then
            Me.Invoke(Sub() _server_RequestDefaultMenu())
            Return
        End If
        ShowDefaultMenu()
    End Sub

    Private Sub ShowDefaultMenu()
        MenuStrip1.Items.Clear()

        ' XenDesk (Erster Button -> Bold)
        Dim finder = New ToolStripMenuItem("XenDesk")
        finder.Font = adamBoldFont

        ' Untermenü für XenDesk
        Dim subItem = New ToolStripMenuItem("Über dieses System")
        subItem.Font = adamMediumFont
        subItem.Tag = "ABOUT_SYSTEM"
        AddHandler subItem.Click, AddressOf OnMenuItemClick
        finder.DropDownItems.Add(subItem)

        Dim subItem2 = New ToolStripMenuItem("Quit XenDesk")
        subItem2.Font = adamMediumFont
        subItem2.Tag = "CLOSE"
        AddHandler subItem2.Click, AddressOf OnMenuItemClick
        finder.DropDownItems.Add(subItem2)

        Dim FileOption = New ToolStripMenuItem("File")
        FileOption.Font = adamMediumFont
        FileOption.Tag = "FILE_MENU"
        AddHandler FileOption.Click, AddressOf OnMenuItemClick


        Dim EditOption = New ToolStripMenuItem("Edit")
        EditOption.Font = adamMediumFont

        Dim DesignOption = New ToolStripMenuItem("Design")
        DesignOption.Font = adamMediumFont

        Dim WindowOption = New ToolStripMenuItem("Window")
        WindowOption.Font = adamMediumFont

        Dim QuickToolsOption = New ToolStripMenuItem("Quick Tools")
        QuickToolsOption.Font = adamMediumFont

        Dim HelpOption = New ToolStripMenuItem("?")
        HelpOption.Font = adamMediumFont


        MenuStrip1.Items.Add(finder)
        MenuStrip1.Items.Add(FileOption)
        MenuStrip1.Items.Add(EditOption)
        MenuStrip1.Items.Add(DesignOption)
        MenuStrip1.Items.Add(WindowOption)
        MenuStrip1.Items.Add(QuickToolsOption)
        MenuStrip1.Items.Add(HelpOption)
    End Sub

    ' EINZIGE BuildMenuRecursive Funktion (zusammengefasst)
    Private Function BuildMenuRecursive(data As MenuItemData, isTopLevel As Boolean, Optional forceBold As Boolean = False) As ToolStripMenuItem
        Dim item As New ToolStripMenuItem(data.Text)
        item.Tag = data.ID

        ' Schrift-Logik
        If isTopLevel AndAlso forceBold Then
            item.Font = adamBoldFont
        Else
            item.Font = adamMediumFont
        End If

        AddHandler item.DropDownOpening, Sub(s, e)
                                             Dim menu = DirectCast(s, ToolStripMenuItem)
                                             SetWindowNoActivate(menu.DropDown.Handle)
                                         End Sub

        If data.SubItems IsNot Nothing AndAlso data.SubItems.Count > 0 Then
            For Each subData In data.SubItems
                item.DropDownItems.Add(BuildMenuRecursive(subData, False, False))
            Next
        Else
            ' Nur Items ohne Untermenü brauchen ein Klick-Event
            AddHandler item.Click, AddressOf OnMenuItemClick
        End If
        Return item
    End Function

    Private Sub OnMenuItemClick(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)
        If item.Tag IsNot Nothing Then
            Dim tagValue As String = item.Tag.ToString()

            _server.SendClickToActiveApp(tagValue)

            Debug.WriteLine("Geklicktes Item: " & item.Text & " mit Tag: " & tagValue)

            If tagValue = "CLOSE" Then
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub CloseAllMenus()
        If MenuStrip1.InvokeRequired Then
            Me.Invoke(Sub() CloseAllMenus())
            Return
        End If

        For Each item As ToolStripItem In MenuStrip1.Items
            If TypeOf item Is ToolStripMenuItem Then
                DirectCast(item, ToolStripMenuItem).DropDown.Close()
            End If
        Next
    End Sub

    Private Sub SetWindowNoActivate(hWnd As IntPtr)
        Const GWL_EXSTYLE As Integer = -20
        Const WS_EX_NOACTIVATE As Integer = &H8000000
        Dim exStyle As Integer = GetWindowLong(hWnd, GWL_EXSTYLE)
        SetWindowLong(hWnd, GWL_EXSTYLE, exStyle Or WS_EX_NOACTIVATE)
    End Sub
End Class