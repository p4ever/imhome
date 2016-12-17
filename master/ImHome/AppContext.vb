Option Explicit On
Imports System.Net
Imports System.Net.NetworkInformation
Public Class AppContext
    Inherits ApplicationContext
    Public options As PopupForm
    Private sRemoteMACAddress
    Private Const No_ERROR = 0
    Dim host As String = New String("")
    Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Integer
    Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIp As Integer, ByVal ScrIP As Integer, ByRef pMacAddr As Long, ByRef PhyAddrLen As Integer) As Integer
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Byte, ByRef src As Long, ByVal bcount As Integer)
    Dim f As New System.Threading.Thread(AddressOf bwmonitor)
    Dim t As New System.Threading.Thread(AddressOf dothemagic85)
    Dim t2 As New System.Threading.Thread(AddressOf dothemagic170)
    Dim t3 As New System.Threading.Thread(AddressOf dothemagic254)
    Public Shared ip As String
    Private MAC As String
    Private found As Boolean = False
    Public Shared tm As New Timer
    Private programOn As Boolean = False
    Private Shared program As New Process
    Public Shared currentrun As String = "Mipony"
    Public Shared stopkill As Boolean = False
#Region " Storage "

    Private WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuDisplayForm As ToolStripMenuItem
    Private WithEvents mnuSep1 As ToolStripSeparator
    Private WithEvents mnuExit As ToolStripMenuItem

#End Region

    Private Sub dothemagic85()
        Dim proc As New Process
        proc.StartInfo.FileName = "cmd.exe"
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.UseShellExecute = False
        Dim i As Integer = 1
        proc.StartInfo.RedirectStandardOutput = True
        While i < 86 And Not found
            proc.StartInfo.Arguments = "/C ping 192.168.1." & i & " -w 2 -n 1"
            proc.Start()
            proc.WaitForExit()
            Dim st As String = proc.StandardOutput.ReadToEnd
            If st.IndexOf("sc") = -1 Then
                Dim dwRemoteIP As Integer
                Dim pMacAddr As Long
                Dim bpMacAddr() As Byte
                Dim PhyAddrLen As Integer
                'convert the string IP into
                'an unsigned long value containing
                'a suitable binary representation
                'of the Internet address given
                dwRemoteIP = inet_addr("192.168.1." & i)
                If dwRemoteIP <> 0 Then
                    'must set this up first!
                    PhyAddrLen = 6

                    'retrieve the remote MAC address
                    If SendARP(dwRemoteIP, 0, pMacAddr, PhyAddrLen) = No_ERROR Then

                        If (pMacAddr <> 0) And (PhyAddrLen <> 0) Then

                            'returned value is a long pointer
                            'to the MAC address, so copy data
                            'to a byte array
                            ReDim bpMacAddr(0 To PhyAddrLen - 1)
                            CopyMemory(bpMacAddr(0), pMacAddr, PhyAddrLen)

                            'convert the byte array to a string 
                            'and return success 
                            sRemoteMACAddress = MakeMacAddress(bpMacAddr, ":")
                            If sRemoteMACAddress.ToString.IndexOf(MAC) <> -1 Then
                                found = True
                            End If
                        End If 'pMacAddr
                        ReDim bpMacAddr(0)
                    End If  'SendARP
                End If  'dwRemoteIP
            End If
            i += 1
        End While
    End Sub

    Private Sub dothemagic170()
        Dim proc As New Process
        proc.StartInfo.FileName = "cmd.exe"
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        Dim i As Integer = 86
        While i < 171 And Not found
            proc.StartInfo.Arguments = "/C ping 192.168.1." & i & " -w 2 -n 1"
            proc.Start()
            proc.WaitForExit()
            Dim st As String = proc.StandardOutput.ReadToEnd
            If st.IndexOf("sc") = -1 Then
                Dim dwRemoteIP As Integer
                Dim pMacAddr As Long
                Dim bpMacAddr() As Byte
                Dim PhyAddrLen As Integer
                'convert the string IP into
                'an unsigned long value containing
                'a suitable binary representation
                'of the Internet address given
                dwRemoteIP = inet_addr("192.168.1." & i)
                If dwRemoteIP <> 0 Then
                    'must set this up first!
                    PhyAddrLen = 6

                    'retrieve the remote MAC address
                    If SendARP(dwRemoteIP, 0, pMacAddr, PhyAddrLen) = No_ERROR Then

                        If (pMacAddr <> 0) And (PhyAddrLen <> 0) Then

                            'returned value is a long pointer
                            'to the MAC address, so copy data
                            'to a byte array
                            ReDim bpMacAddr(0 To PhyAddrLen - 1)
                            CopyMemory(bpMacAddr(0), pMacAddr, PhyAddrLen)

                            'convert the byte array to a string 
                            'and return success 
                            sRemoteMACAddress = MakeMacAddress(bpMacAddr, ":")
                            If sRemoteMACAddress.ToString.IndexOf(MAC) <> -1 Then
                                found = True
                            End If
                        End If 'pMacAddr
                        ReDim bpMacAddr(0)
                    End If  'SendARP
                End If  'dwRemoteIP
            End If
            i += 1
        End While
    End Sub

    Private Sub dothemagic254()
        Dim proc As New Process
        proc.StartInfo.FileName = "cmd.exe"
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        Dim i As Integer = 171
        While i < 255 And Not found
            proc.StartInfo.Arguments = "/C ping 192.168.1." & i & " -w 2 -n 1"
            proc.Start()
            proc.WaitForExit()
            Dim st As String = proc.StandardOutput.ReadToEnd
            If st.IndexOf("sc") = -1 Then
                Dim dwRemoteIP As Integer
                Dim pMacAddr As Long
                Dim bpMacAddr() As Byte
                Dim PhyAddrLen As Integer
                'convert the string IP into
                'an unsigned long value containing
                'a suitable binary representation
                'of the Internet address given
                dwRemoteIP = inet_addr("192.168.1." & i)
                If dwRemoteIP <> 0 Then
                    'must set this up first!
                    PhyAddrLen = 6

                    'retrieve the remote MAC address
                    If SendARP(dwRemoteIP, 0, pMacAddr, PhyAddrLen) = No_ERROR Then

                        If (pMacAddr <> 0) And (PhyAddrLen <> 0) Then

                            'returned value is a long pointer
                            'to the MAC address, so copy data
                            'to a byte array
                            ReDim bpMacAddr(0 To PhyAddrLen - 1)
                            CopyMemory(bpMacAddr(0), pMacAddr, PhyAddrLen)

                            'convert the byte array to a string 
                            'and return success 
                            sRemoteMACAddress = MakeMacAddress(bpMacAddr, ":")
                            If sRemoteMACAddress.ToString.IndexOf(MAC) <> -1 Then
                                found = True
                            End If
                        End If 'pMacAddr
                        ReDim bpMacAddr(0)
                    End If  'SendARP
                End If  'dwRemoteIP
            End If
            i += 1
        End While
    End Sub

    Private Sub MyTickHandler(ByVal sender As Object, ByVal e As EventArgs)
        If Not t.IsAlive And Not t2.IsAlive And Not t3.IsAlive Then
            Console.WriteLine("end")
            Dim manSearch As Boolean = False
            'check if the program has already opened by the user
            If Process.GetProcessesByName("Mipony").Length > 0 Or Process.GetProcessesByName("uTorrent").Length > 0 Then
                manSearch = True
            End If
            If Not manSearch Then
                If found Then
                    Try
                        programOn = False
                        Dim prg = Process.GetProcessesByName(currentrun)
                        For j As Integer = 0 To prg.Length - 1
                            '  prg(j).CloseMainWindow()
                            '  prg(j).Close()
                            prg(j).Kill()
                        Next
                    Catch ex As Exception

                    End Try
                    '   tm.Interval = 300000
                Else
                    If Not programOn Then
                        program.Start()
                        f.Start()
                        currentrun = "Mipony"
                        programOn = True
                    End If
                    tm.Interval = 50000
                End If
                found = False
                f = New System.Threading.Thread(AddressOf bwmonitor)
                t = New System.Threading.Thread(AddressOf dothemagic85)
                t2 = New System.Threading.Thread(AddressOf dothemagic170)
                t3 = New System.Threading.Thread(AddressOf dothemagic254)
                t.Start()
                t2.Start()
                t3.Start()
            End If
        End If
    End Sub


    Private Function MakeMacAddress(b() As Byte, sDelim As String) As String

        Dim cnt As Long
        Dim buff As String


        'so far, MAC addresses are
        'exactly 6 segments in size (0-5)
        If UBound(b) = 5 Then

            'concatenate the first five values
            'together and separate with the
            'delimiter char
            For cnt = 0 To 4
                buff = buff & Right$("00" & Hex(b(cnt)), 2) & sDelim
            Next

            'and append the last value
            buff = buff & Right$("00" & Hex(b(5)), 2)

        End If  'UBound(b)

        MakeMacAddress = buff

    End Function

#Region " Constructor "

    Public Sub New()
        'Initialize the menus
        mnuDisplayForm = New ToolStripMenuItem("Display form")
        mnuSep1 = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Exit")
        MainMenu = New ContextMenuStrip
        MainMenu.Items.AddRange(New ToolStripItem() {mnuDisplayForm, mnuSep1, mnuExit})
        'Initialize the tray
        Tray = New NotifyIcon
        Tray.Icon = My.Resources.TrayIcon
        Tray.ContextMenuStrip = MainMenu
        Tray.Text = "Formless tray application"
        host = Dns.GetHostName()
        tm.Interval = 50000
        AddHandler tm.Tick, AddressOf MyTickHandler
        tm.Start()
        'Display
        Tray.Visible = True
        'MAC = "B4:8B:19:6D:2F:D9"
        MAC = "6C:40:08:90:98:EE"
        program.StartInfo.FileName = "C:\Program Files (x86)\MiPony\MiPony.exe"
        Console.WriteLine("start")
        t.Start()
        t2.Start()
        t3.Start()
    End Sub

#End Region

    Sub bwmonitor()
        Dim pc As New PerformanceCounterCategory("Network Interface")
        Dim instance As String = pc.GetInstanceNames(2)
        Dim br As New PerformanceCounter("Network Interface", "Bytes Received/Sec", instance)
        Dim k As Integer = 0
        Dim kbRecieved As Integer = 0
        Do
            k = k + 1
            kbRecieved += br.NextValue / 1024
            Threading.Thread.Sleep(1000)
        Loop Until k = 30
        If kbRecieved / k < 240 Then
            ' change program
            programOn = False
            Dim prg = Process.GetProcessesByName("Mipony")
            For j As Integer = 0 To prg.Length - 1
                prg(j).CloseMainWindow()
            Next
            If Not programOn Then
                program.StartInfo.FileName = "C:\Users\Alby\AppData\Roaming\uTorrent\uTorrent.exe"
                program.Start()
                programOn = True
                currentrun = "uTorrent"
                program.StartInfo.FileName = "C:\Program Files (x86)\MiPony\MiPony.exe"
            End If
        End If
        br.Dispose()
    End Sub

#Region " Event handlers "

    Private Sub AppContext_ThreadExit(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Me.ThreadExit
        'Guarantees that the icon will not linger.
        Tray.Visible = False
    End Sub

#End Region

    Private Sub mnuDisplayForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles mnuDisplayForm.Click
        options = New PopupForm
        options.ShowDialog()
    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles mnuExit.Click
        ExitApplication()
    End Sub

    Private Sub Tray_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles Tray.DoubleClick
        Dim options = New PopupForm
    End Sub
End Class
