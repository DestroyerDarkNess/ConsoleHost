Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Win32.SafeHandles

Public Class ConsoleHost : Inherits Panel

#Region " Pinvoke "

    <DllImport("kernel32.dll")>
    Private Shared Function AllocConsole() As Boolean
    End Function

    <DllImport("kernel32", SetLastError:=True)>
    Private Shared Function AttachConsole(ByVal dwProcessId As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetConsoleWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, <Out> ByRef lpdwProcessId As Integer) As UInteger
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function FreeConsole() As Boolean
    End Function

    <DllImport("kernel32.dll", EntryPoint:="GetStdHandle", SetLastError:=True, CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function GetStdHandle(ByVal nStdHandle As Integer) As IntPtr
    End Function

    'SetParent

    <DllImport("user32.dll", EntryPoint:="SetParent")>
    Private Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowPos")>
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(ByVal hWnd As System.IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function

#End Region

#Region " Fix White Bar Console "

    <DllImport("user32.dll")>
    Private Shared Function UpdateWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SendMessage(hWnd As IntPtr, wMsg As UInteger, wParam As UIntPtr, lParam As IntPtr) As Integer
    End Function

    <DllImport("user32")>
    Public Shared Function ShowScrollBar(ByVal hWnd As System.IntPtr, ByVal wBar As Integer, ByVal bShow As Boolean) As Boolean
    End Function

    Public Function UpdateWindowEx()
        ' SendMessage(HandleEx, &HB6, 0, 20)
        ShowScrollBar(HandleEx, 1, True)
        Return UpdateWindow(HandleEx)
    End Function

#End Region

#Region " Properties "

    Private HandleEx As IntPtr = IntPtr.Zero
    Public ReadOnly Property ConsoleHandle As IntPtr
        <DebuggerStepThrough>
        Get
            Return Me.HandleEx
        End Get
    End Property

    Private ProcessEx As Process = Nothing
    Public ReadOnly Property TargetProcess As Process
        <DebuggerStepThrough>
        Get
            Return Me.ProcessEx
        End Get
    End Property

    Private ConsoleReadyEx As Boolean = False
    Public ReadOnly Property ConsoleReady As Boolean
        <DebuggerStepThrough>
        Get
            Return Me.ConsoleReadyEx
        End Get
    End Property

    Private ManualWriterEx As StreamWriter = Nothing
    Public ReadOnly Property ManualWriter As StreamWriter
        <DebuggerStepThrough>
        Get
            Return Me.ManualWriterEx
        End Get
    End Property

#End Region

#Region " Declare "

    Private Const SW_SHOWNOACTIVATE As Integer = 4
    Private Const HWND_BOTTOM As Integer = &H1

    Private Const STD_OUTPUT_HANDLE As Integer = -11
    Private Const MY_CODE_PAGE As Integer = 437

#End Region

#Region " Constructor "

    Public Sub New()
        Me.BackColor = Color.Black
        Me.ForeColor = Color.White
    End Sub

    Public Function Initialize(Optional ByVal TargetProcess As Process = Nothing)
        ProcessEx = TargetProcess
        Dim Loaded As Boolean = False

        If ProcessEx Is Nothing Then

            AllocConsole()

            Dim hwnd As IntPtr = GetConsoleWindow()

            Debug.WriteLine("GetConsoleWindow = " & hwnd.ToString)

            If hwnd = IntPtr.Zero Then

                Dim Proc As New System.Diagnostics.Process
                Proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                Proc.StartInfo.CreateNoWindow = False
                Proc.StartInfo.UseShellExecute = True
                Proc.StartInfo.FileName = "cmd.exe"
                Proc.StartInfo.Arguments = "/k @echo off & Title ConsoleHost & cls & pause >nul"
                Proc.Start()
                AttachConsole(Proc.Id)
                Loaded = SetByProcess(Proc)

                HandleEx = Proc.MainWindowHandle
                ProcessEx = Proc

                Debug.WriteLine("GetConsoleWindow Failed!")
                Debug.WriteLine("Helper CMD Host Loaded ")

            Else
                Debug.WriteLine("GetConsoleWindow Has been loaded!")
                HandleEx = hwnd
                Loaded = SetByHandle(hwnd)
            End If
        Else

            HandleEx = TargetProcess.MainWindowHandle
            AttachConsole(TargetProcess.Id)
            Loaded = SetByProcess(TargetProcess)

        End If

        Dim Asynctask As New Task(New Action(Async Sub()
                                                 Dim Seconds As Integer = 1
                                                 System.Threading.Thread.Sleep(500)
                                                 UpdateWindowEx()
                                             End Sub), TaskCreationOptions.PreferFairness)
        Asynctask.Start()

        Return Loaded
    End Function

    Public Function Unsecure_Initialize() As Boolean
        Try

            AllocConsole()

            Dim hwnd As IntPtr = GetConsoleWindow()
            HandleEx = hwnd

            Dim stdHandle As IntPtr = GetStdHandle(STD_OUTPUT_HANDLE)
            Dim safeFileHandle As SafeFileHandle = New SafeFileHandle(stdHandle, True)
            Dim fileStream As FileStream = New FileStream(safeFileHandle, FileAccess.Write)
            Dim encoding As Encoding = System.Text.Encoding.GetEncoding(MY_CODE_PAGE)
            Dim standardOutput As StreamWriter = New StreamWriter(fileStream, encoding)
            standardOutput.AutoFlush = True

            Console.SetOut(standardOutput)
            ManualWriterEx = standardOutput
            Debug.WriteLine("GetConsoleWindow = " & hwnd.ToString)

            Dim Embed As Boolean = SetByHandle(hwnd)

            Dim Asynctask As New Task(New Action(Async Sub()
                                                     Dim Seconds As Integer = 1
                                                     System.Threading.Thread.Sleep(500)
                                                     UpdateWindowEx()
                                                 End Sub), TaskCreationOptions.PreferFairness)
            Asynctask.Start()

            Return Embed
        Catch ex As Exception
            Debug.WriteLine("Unsecure_Initialize = " & ex.Message.ToString)
            Return False
        End Try
    End Function

#End Region

#Region " Private Methods "

    Private Function SetByHandle(ByVal Proc_MainWindowHandle As IntPtr)
        Try
            Dim SetCorrectParent As Boolean = False
            Dim Procede As Boolean = False
            Dim Limit As Integer = 0
            For i As Integer = 0 To 2

                Dim SetParentHandle As IntPtr = SetParent(Proc_MainWindowHandle, Me.Handle)
                Debug.WriteLine("SetParentHandle = " & SetParentHandle.ToString)

                If SetParentHandle <> IntPtr.Zero Then
                    SetWindowPos(Proc_MainWindowHandle, New IntPtr(HWND_BOTTOM), 0, 0, 0, 0, 1)
                    SetCorrectParent = True
                End If

                If SetCorrectParent = True Then

                    Dim placement As WINDOWPLACEMENT = GetPlacement(Proc_MainWindowHandle)
                    Debug.WriteLine("placement = " & placement.showCmd.ToString)
                    If placement.showCmd.ToString = "Normal" Then
                        Dim FakeFullSc As Boolean = FullScreenEmulation(Proc_MainWindowHandle)
                        Procede = True
                    End If

                    If Procede = True Then
                        Limit += 1
                        Debug.WriteLine(" Limit = " & Limit.ToString)
                        If placement.showCmd.ToString = "Maximized" Then
                            Dim FakeFullSc As Boolean = FullScreenEmulation(Proc_MainWindowHandle)
                            ShowWindow(Proc_MainWindowHandle, SW_SHOWNOACTIVATE)
                            Exit For
                        ElseIf Limit = 2 Then
                            Exit For
                        End If
                    End If
                End If

                System.Windows.Forms.Application.DoEvents()

                i -= 1
            Next

            Console.Clear()

            Debug.WriteLine("CosoleHost Loaded!")
            Return True
        Catch ex As Exception
            Debug.WriteLine("CosoleHost Error: " & ex.Message)
            Return False
        End Try
    End Function



    Private Function SetByProcess(ByVal Proc As Process)
        Try
            Dim SetCorrectParent As Boolean = False
            Dim Procede As Boolean = False
            Dim Limit As Integer = 0
            For i As Integer = 0 To 2
                Dim SetParentHandle As IntPtr = SetParent(Proc.MainWindowHandle, Me.Handle)
                Debug.WriteLine("SetParentHandle = " & SetParentHandle.ToString)

                If SetParentHandle <> IntPtr.Zero Then
                    SetWindowPos(Proc.MainWindowHandle, New IntPtr(HWND_BOTTOM), 0, 0, 0, 0, 1)
                    SetCorrectParent = True
                End If

                If SetCorrectParent = True Then
                    Dim placement As WINDOWPLACEMENT = GetPlacement(Proc.MainWindowHandle)
                    Debug.WriteLine("placement = " & placement.showCmd.ToString)
                    If placement.showCmd.ToString = "Normal" Then
                        Dim FakeFullSc As Boolean = FullScreenEmulation(Proc.MainWindowHandle)
                        Procede = True
                    End If

                    If Procede = True Then
                        Limit += 1
                        Debug.WriteLine(" Limit = " & Limit.ToString)
                        If placement.showCmd.ToString = "Maximized" Then
                            Dim FakeFullSc As Boolean = FullScreenEmulation(Proc.MainWindowHandle)
                            ShowWindow(Proc.MainWindowHandle, SW_SHOWNOACTIVATE)
                            Exit For
                        ElseIf Limit = 2 Then
                            Exit For
                        End If
                    End If
                End If
                System.Windows.Forms.Application.DoEvents()
                i -= 1
            Next

            Debug.WriteLine("CosoleHost Loaded!")

            Return True
        Catch ex As Exception
            Debug.WriteLine("CosoleHost Error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function FullScreenEmulation(ByVal Proc_MainWindowHandle As IntPtr) As Boolean
        Try
            Dim HWND As IntPtr = Proc_MainWindowHandle
            For i As Integer = 0 To 2
                Win32.Helpers.SetWindowStyle.SetWindowStyle(HWND, Win32.Helpers.SetWindowStyle.WindowStyles.WS_BORDER)
                Win32.Helpers.SetWindowState.SetWindowState(HWND, Win32.Helpers.SetWindowState.WindowState.Maximize)
            Next
            BringMainWindowToFront(HWND)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region " Set Focus "

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hwnd As IntPtr) As Integer
    End Function

    Private Shared FisrsFocus As Boolean = False

    Private Sub BringMainWindowToFront(ByVal Proc_MainWindowHandle As IntPtr)
        If FisrsFocus = False Then
            SetForegroundWindow(Proc_MainWindowHandle)
            FisrsFocus = True
        End If
    End Sub

    Public Sub ActiveConsoleFocus()
        SetForegroundWindow(HandleEx)
    End Sub

#End Region

#Region " Check FakeFullscreen "

    Private Shared Function GetPlacement(ByVal hwnd As IntPtr) As WINDOWPLACEMENT
        Dim placement As WINDOWPLACEMENT = New WINDOWPLACEMENT()
        placement.length = Marshal.SizeOf(placement)
        GetWindowPlacement(hwnd, placement)
        Return placement
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    <Serializable>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure WINDOWPLACEMENT
        Public length As Integer
        Public flags As Integer
        Public showCmd As ShowWindowCommands
        Public ptMinPosition As System.Drawing.Point
        Public ptMaxPosition As System.Drawing.Point
        Public rcNormalPosition As System.Drawing.Rectangle
    End Structure

    Friend Enum ShowWindowCommands As Integer
        Hide = 0
        Normal = 1
        Minimized = 2
        Maximized = 3
    End Enum

#End Region

End Class

Namespace Win32.Helpers

    ' ***********************************************************************
    ' Author   : Elektro
    ' Modified : 02-21-2014
    ' ***********************************************************************
    ' <copyright file="SendInputs.vb" company="Elektro Studios">
    '     Copyright (c) Elektro Studios. All rights reserved.
    ' </copyright>
    ' ***********************************************************************

#Region " Usage Examples "

    'Private Sub Test() Handles Button1.Click

    ' AppActivate(Process.GetProcessesByName("notepad").First.Id)

    ' Dim c As Char = Convert.ToChar(Keys.Oemtilde) ' Ñ
    ' Dim Result As Integer = SendInputs.SendKey(Convert.ToChar(c.ToString.ToLower))
    ' MessageBox.Show(String.Format("Successfull events: {0}", CStr(Result)))

    ' SendInputs.SendKey(Keys.Enter)
    ' SendInputs.SendKey(Convert.ToChar(Keys.Back))
    ' SendInputs.SendKeys("Hello World", True)
    ' SendInputs.SendKey(Convert.ToChar(Keys.D0))
    ' SendInputs.SendKeys(Keys.Insert, BlockInput:=True)

    ' SendInputs.MouseClick(SendInputs.MouseButton.RightPress, False)
    ' SendInputs.MouseMove(5, -5)
    ' SendInputs.MousePosition(New Point(100, 500))

    'End Sub

#End Region

    ''' <summary>
    ''' Synthesizes keystrokes, mouse motions, and button clicks.
    ''' </summary>
    Public Class SendInputs

#Region " P/Invoke "

        Friend Class NativeMethods

#Region " Methods "

            ''' <summary>
            ''' Blocks keyboard and mouse input events from reaching applications.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646290%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="fBlockIt">
            ''' The function's purpose. 
            ''' If this parameter is 'TRUE', keyboard and mouse input events are blocked. 
            ''' If this parameter is 'FALSE', keyboard and mouse events are unblocked. 
            ''' </param>
            ''' <returns>
            ''' If the function succeeds, the return value is nonzero.
            ''' If input is already blocked, the return value is zero.
            ''' </returns>
            ''' <remarks>
            ''' Note that only the thread that blocked input can successfully unblock input.
            ''' </remarks>
            <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall,
        SetLastError:=True)>
            Friend Shared Function BlockInput(
               ByVal fBlockIt As Boolean
        ) As Integer
            End Function

            ''' <summary>
            ''' Synthesizes keystrokes, mouse motions, and button clicks.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646310%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="nInputs">
            ''' Indicates the number of structures in the pInputs array.
            ''' </param>
            ''' <param name="pInputs">
            ''' Indicates an Array of 'INPUT' structures.
            ''' Each structure represents an event to be inserted into the keyboard or mouse input stream.
            ''' </param>
            ''' <param name="cbSize">
            ''' The size, in bytes, of an 'INPUT' structure.
            ''' If 'cbSize' is not the size of an 'INPUT' structure, the function fails.
            ''' </param>
            ''' <returns>
            ''' The function returns the number of events that it successfully 
            ''' inserted into the keyboard or mouse input stream. 
            ''' If the function returns zero, the input was already blocked by another thread.
            ''' </returns>
            <DllImport("user32.dll", SetLastError:=True)>
            Friend Shared Function SendInput(
               ByVal nInputs As Integer,
               <MarshalAs(UnmanagedType.LPArray), [In]> ByVal pInputs As Input(),
               ByVal cbSize As Integer
        ) As Integer
            End Function

#End Region

#Region " Enumerations "

            ''' <summary>
            ''' VirtualKey codes.
            ''' </summary>
            Friend Enum VirtualKeys As Short

                ''' <summary>
                ''' The Shift key.
                ''' VK_SHIFT
                ''' </summary>
                SHIFT = &H10S

                ''' <summary>
                ''' The DEL key.
                ''' VK_DELETE
                ''' </summary>
                DELETE = 46S

                ''' <summary>
                ''' The ENTER key.
                ''' VK_RETURN
                ''' </summary>
                [RETURN] = 13S

            End Enum

            ''' <summary>
            ''' The type of the input event.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646270%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Enumeration used for 'type' parameter of 'INPUT' structure")>
            Friend Enum InputType As Integer

                ''' <summary>
                ''' The event is a mouse event.
                ''' Use the mi structure of the union.
                ''' </summary>
                Mouse = 0

                ''' <summary>
                ''' The event is a keyboard event.
                ''' Use the ki structure of the union.
                ''' </summary>
                Keyboard = 1

                ''' <summary>
                ''' The event is a hardware event.
                ''' Use the hi structure of the union.
                ''' </summary>
                Hardware = 2

            End Enum

            ''' <summary>
            ''' Specifies various aspects of a keystroke. 
            ''' This member can be certain combinations of the following values. 
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646271%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Enumeration used for 'dwFlags' parameter of 'KeyboardInput' structure")>
            <Flags>
            Friend Enum KeyboardInput_Flags As Integer

                ''' <summary>
                ''' If specified, the scan code was preceded by a prefix byte that has the value '0xE0' (224).
                ''' </summary>
                ExtendedKey = &H1

                ''' <summary>
                ''' If specified, the key is being pressed.
                ''' </summary>
                KeyDown = &H0

                ''' <summary>
                ''' If specified, the key is being released. 
                ''' If not specified, the key is being pressed.
                ''' </summary>
                KeyUp = &H2

                ''' <summary>
                ''' If specified, 'wScan' identifies the key and 'wVk' is ignored. 
                ''' </summary>
                ScanCode = &H8

                ''' <summary>
                ''' If specified, the system synthesizes a 'VK_PACKET' keystroke. 
                ''' The 'wVk' parameter must be '0'. 
                ''' This flag can only be combined with the 'KEYEVENTF_KEYUP' flag. 
                ''' </summary>
                Unicode = &H4

            End Enum

            ''' <summary>
            ''' A set of bit flags that specify various aspects of mouse motion and button clicks. 
            ''' The bits in this member can be any reasonable combination of the following values. 
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646273%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Enumeration used for 'dwFlags' parameter of 'MouseInput' structure")>
            <Flags>
            Friend Enum MouseInput_Flags As Integer

                ''' <summary>
                ''' The 'dx' and 'dy' members contain normalized absolute coordinates. 
                ''' If the flag is not set, 'dx' and 'dy' contain relative data 
                ''' (the change in position since the last reported position). 
                ''' This flag can be set, or not set, 
                ''' regardless of what kind of mouse or other pointing device, if any, is connected to the system. 
                ''' </summary>
                Absolute = &H8000I

                ''' <summary>
                ''' Movement occurred.
                ''' </summary>
                Move = &H1I

                ''' <summary>
                ''' The 'WM_MOUSEMOVE' messages will not be coalesced. 
                ''' The default behavior is to coalesce 'WM_MOUSEMOVE' messages. 
                ''' </summary>
                Move_NoCoalesce = &H2000I

                ''' <summary>
                ''' The left button was pressed.
                ''' </summary>
                LeftDown = &H2I

                ''' <summary>
                ''' The left button was released.
                ''' </summary>
                LeftUp = &H4I

                ''' <summary>
                ''' The right button was pressed.
                ''' </summary>
                RightDown = &H8I

                ''' <summary>
                ''' The right button was released.
                ''' </summary>
                RightUp = &H10I

                ''' <summary>
                ''' The middle button was pressed.
                ''' </summary>
                MiddleDown = &H20I

                ''' <summary>
                ''' The middle button was released.
                ''' </summary>
                MiddleUp = &H40I

                ''' <summary>
                ''' Maps coordinates to the entire desktop. 
                ''' Must be used in combination with 'Absolute'.
                ''' </summary>
                VirtualDesk = &H4000I

                ''' <summary>
                ''' The wheel was moved, if the mouse has a wheel. 
                ''' The amount of movement is specified in 'mouseData'. 
                ''' </summary>
                Wheel = &H800I

                ''' <summary>
                ''' The wheel was moved horizontally, if the mouse has a wheel. 
                ''' The amount of movement is specified in 'mouseData'. 
                ''' </summary>
                HWheel = &H1000I

                ''' <summary>
                ''' An X button was pressed.
                ''' </summary>
                XDown = &H80I

                ''' <summary>
                ''' An X button was released.
                ''' </summary>
                XUp = &H100I

            End Enum

#End Region

#Region " Structures "

            ''' <summary>
            ''' Used by 'SendInput' function
            ''' to store information for synthesizing input events such as keystrokes, mouse movement, and mouse clicks.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646270%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Structure used for 'INPUT' parameter of 'SendInput' API method")>
            <StructLayout(LayoutKind.Explicit)>
            Friend Structure Input

                ' ******
                '  NOTE
                ' ******
                ' Field offset for 32 bit machine: 4
                ' Field offset for 64 bit machine: 8

                ''' <summary>
                ''' The type of the input event.
                ''' </summary>
                <FieldOffset(0)>
                Public type As InputType

                ''' <summary>
                ''' The information about a simulated mouse event.
                ''' </summary>
                <FieldOffset(8)>
                Public mi As MouseInput

                ''' <summary>
                ''' The information about a simulated keyboard event.
                ''' </summary>
                <FieldOffset(8)>
                Public ki As KeyboardInput

                ''' <summary>
                ''' The information about a simulated hardware event.
                ''' </summary>
                <FieldOffset(8)>
                Public hi As HardwareInput

            End Structure

            ''' <summary>
            ''' Contains information about a simulated mouse event.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646273%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Structure used for 'mi' parameter of 'INPUT' structure")>
            Friend Structure MouseInput

                ''' <summary>
                ''' The absolute position of the mouse, 
                ''' or the amount of motion since the last mouse event was generated, 
                ''' depending on the value of the dwFlags member.
                ''' Absolute data is specified as the 'x' coordinate of the mouse; 
                ''' relative data is specified as the number of pixels moved.
                ''' </summary>
                Public dx As Integer

                ''' <summary>
                ''' The absolute position of the mouse, 
                ''' or the amount of motion since the last mouse event was generated, 
                ''' depending on the value of the dwFlags member. 
                ''' Absolute data is specified as the 'y' coordinate of the mouse; 
                ''' relative data is specified as the number of pixels moved. 
                ''' </summary>
                Public dy As Integer

                ''' <summary>
                ''' If 'dwFlags' contains 'MOUSEEVENTF_WHEEL', 
                ''' then 'mouseData' specifies the amount of wheel movement. 
                ''' A positive value indicates that the wheel was rotated forward, away from the user; 
                ''' a negative value indicates that the wheel was rotated backward, toward the user. 
                ''' One wheel click is defined as 'WHEEL_DELTA', which is '120'.
                ''' 
                ''' If 'dwFlags' does not contain 'MOUSEEVENTF_WHEEL', 'MOUSEEVENTF_XDOWN', or 'MOUSEEVENTF_XUP', 
                ''' then mouseData should be '0'. 
                ''' </summary>
                Public mouseData As Integer

                ''' <summary>
                ''' A set of bit flags that specify various aspects of mouse motion and button clicks. 
                ''' The bits in this member can be any reasonable combination of the following values.
                ''' The bit flags that specify mouse button status are set to indicate changes in status, 
                ''' not ongoing conditions. 
                ''' For example, if the left mouse button is pressed and held down, 
                ''' 'MOUSEEVENTF_LEFTDOWN' is set when the left button is first pressed, 
                ''' but not for subsequent motions. 
                ''' Similarly, 'MOUSEEVENTF_LEFTUP' is set only when the button is first released. 
                ''' 
                ''' You cannot specify both the 'MOUSEEVENTF_WHEE'L flag 
                ''' and either 'MOUSEEVENTF_XDOWN' or 'MOUSEEVENTF_XUP' flags simultaneously in the 'dwFlags' parameter, 
                ''' because they both require use of the 'mouseData' field. 
                ''' </summary>
                Public dwFlags As MouseInput_Flags

                ''' <summary>
                ''' The time stamp for the event, in milliseconds. 
                ''' If this parameter is '0', the system will provide its own time stamp. 
                ''' </summary>
                Public time As Integer

                ''' <summary>
                ''' An additional value associated with the mouse event. 
                ''' An application calls 'GetMessageExtraInfo' to obtain this extra information. 
                ''' </summary>
                Public dwExtraInfo As IntPtr

            End Structure

            ''' <summary>
            ''' Contains information about a simulated keyboard event.
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646271%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Structure used for 'ki' parameter of 'INPUT' structure")>
            Friend Structure KeyboardInput

                ''' <summary>
                ''' A virtual-key code. 
                ''' The code must be a value in the range '1' to '254'. 
                ''' If the 'dwFlags' member specifies 'KEYEVENTF_UNICODE', wVk must be '0'. 
                ''' </summary>
                Public wVk As Short

                ''' <summary>
                ''' A hardware scan code for the key. 
                ''' If 'dwFlags' specifies 'KEYEVENTF_UNICODE', 
                ''' 'wScan' specifies a Unicode character which is to be sent to the foreground application. 
                ''' </summary>
                Public wScan As Short

                ''' <summary>
                ''' Specifies various aspects of a keystroke.
                ''' </summary>
                Public dwFlags As KeyboardInput_Flags

                ''' <summary>
                ''' The time stamp for the event, in milliseconds. 
                ''' If this parameter is '0', the system will provide its own time stamp.
                ''' </summary>
                Public time As Integer

                ''' <summary>
                ''' An additional value associated with the keystroke. 
                ''' Use the 'GetMessageExtraInfo' function to obtain this information. 
                ''' </summary>
                Public dwExtraInfo As IntPtr

            End Structure

            ''' <summary>
            ''' Contains information about a simulated message generated by an input device other than a keyboard or mouse. 
            ''' For more info see here:
            ''' http://msdn.microsoft.com/en-us/library/windows/desktop/ms646269%28v=vs.85%29.aspx
            ''' </summary>
            <Description("Structure used for 'hi' parameter of 'INPUT' structure")>
            Friend Structure HardwareInput

                ''' <summary>
                ''' The message generated by the input hardware. 
                ''' </summary>
                Public uMsg As Integer

                ''' <summary>
                ''' The low-order word of the lParam parameter for uMsg. 
                ''' </summary>
                Public wParamL As Short

                ''' <summary>
                ''' The high-order word of the lParam parameter for uMsg. 
                ''' </summary>
                Public wParamH As Short

            End Structure

#End Region

        End Class

#End Region

#Region " Enumerations "

        ''' <summary>
        ''' Indicates a mouse button.
        ''' </summary>
        <Description("Enumeration used for 'MouseAction' parameter of 'MouseClick' function.")>
        Public Enum MouseButton As Integer

            ''' <summary>
            ''' Hold the left button.
            ''' </summary>
            LeftDown = &H2I

            ''' <summary>
            ''' Release the left button.
            ''' </summary>
            LeftUp = &H4I

            ''' <summary>
            ''' Hold the right button.
            ''' </summary>
            RightDown = &H8I

            ''' <summary>
            ''' Release the right button.
            ''' </summary>
            RightUp = &H10I

            ''' <summary>
            ''' Hold the middle button.
            ''' </summary>
            MiddleDown = &H20I

            ''' <summary>
            ''' Release the middle button.
            ''' </summary>
            MiddleUp = &H40I

            ''' <summary>
            ''' Press the left button.
            ''' ( Hold + Release )
            ''' </summary>
            LeftPress = LeftDown + LeftUp

            ''' <summary>
            ''' Press the Right button.
            ''' ( Hold + Release )
            ''' </summary>
            RightPress = RightDown + RightUp

            ''' <summary>
            ''' Press the Middle button.
            ''' ( Hold + Release )
            ''' </summary>
            MiddlePress = MiddleDown + MiddleUp

        End Enum

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Sends a keystroke.
        ''' </summary>
        ''' <param name="key">
        ''' Indicates the keystroke to simulate.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the keystroke is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the keyboard input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function SendKey(ByVal key As Char,
                                   Optional BlockInput As Boolean = False) As Integer

            ' Block Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(True)

            ' The inputs structures to send.
            Dim Inputs As New List(Of NativeMethods.Input)

            ' The current input to add into the Inputs list.
            Dim CurrentInput As New NativeMethods.Input

            ' Determines whether a character is an alphabetic letter.
            Dim IsAlphabetic As Boolean = Not (key.ToString.ToUpper = key.ToString.ToLower)

            ' Determines whether a character is an uppercase alphabetic letter.
            Dim IsUpperCase As Boolean =
            (key.ToString = key.ToString.ToUpper) AndAlso Not (key.ToString.ToUpper = key.ToString.ToLower)

            ' Determines whether the CapsLock key is pressed down.
            Dim CapsLockON As Boolean = My.Computer.Keyboard.CapsLock

            ' Set the passed key to upper-case.
            If IsAlphabetic AndAlso Not IsUpperCase Then
                key = Convert.ToChar(key.ToString.ToUpper)
            End If

            ' If character is alphabetic and is UpperCase and CapsLock is pressed down,
            ' OrElse character is alphabetic and is not UpperCase and CapsLock is not pressed down,
            ' OrElse character is not alphabetic.
            If (IsAlphabetic AndAlso IsUpperCase AndAlso CapsLockON) _
        OrElse (IsAlphabetic AndAlso Not IsUpperCase AndAlso Not CapsLockON) _
        OrElse (Not IsAlphabetic) Then

                ' Hold the character key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = Convert.ToInt16(CChar(key))
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyDown
                End With : Inputs.Add(CurrentInput)

                ' Release the character key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = Convert.ToInt16(CChar(key))
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyUp
                End With : Inputs.Add(CurrentInput)

                ' If character is alphabetic and is UpperCase and CapsLock is not pressed down,
                ' OrElse character is alphabetic and is not UpperCase and CapsLock is pressed down.
            ElseIf (IsAlphabetic AndAlso IsUpperCase AndAlso Not CapsLockON) _
        OrElse (IsAlphabetic AndAlso Not IsUpperCase AndAlso CapsLockON) Then

                ' Hold the Shift key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = NativeMethods.VirtualKeys.SHIFT
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyDown
                End With : Inputs.Add(CurrentInput)

                ' Hold the character key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = Convert.ToInt16(CChar(key))
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyDown
                End With : Inputs.Add(CurrentInput)

                ' Release the character key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = Convert.ToInt16(CChar(key))
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyUp
                End With : Inputs.Add(CurrentInput)

                ' Release the Shift key.
                With CurrentInput
                    .type = NativeMethods.InputType.Keyboard
                    .ki.wVk = NativeMethods.VirtualKeys.SHIFT
                    .ki.dwFlags = NativeMethods.KeyboardInput_Flags.KeyUp
                End With : Inputs.Add(CurrentInput)

            End If ' UpperCase And My.Computer.Keyboard.CapsLock is...

            ' Send the input key.
            Return NativeMethods.SendInput(Inputs.Count, Inputs.ToArray,
                                       Marshal.SizeOf(GetType(NativeMethods.Input)))

            ' Unblock Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(False)

        End Function

        ''' <summary>
        ''' Sends a keystroke.
        ''' </summary>
        ''' <param name="key">
        ''' Indicates the keystroke to simulate.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the keystroke is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the keyboard input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function SendKey(ByVal key As Keys,
                                   Optional BlockInput As Boolean = False) As Integer

            Return SendKey(Convert.ToChar(key), BlockInput)

        End Function

        ''' <summary>
        ''' Sends a string.
        ''' </summary>
        ''' <param name="String">
        ''' Indicates the string to send.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the keystroke is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the keyboard input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function SendKeys(ByVal [String] As String,
                                    Optional BlockInput As Boolean = False) As Integer

            Dim SuccessCount As Integer = 0

            ' Block Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(True)

            For Each c As Char In [String]
                SuccessCount += SendKey(c, BlockInput:=False)
            Next c

            ' Unblock Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(False)

            Return SuccessCount

        End Function

        ''' <summary>
        ''' Slices the mouse position.
        ''' </summary>
        ''' <param name="Offset">
        ''' Indicates the offset, in coordinates.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the mouse movement is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the mouse input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function MouseMove(ByVal Offset As Point,
                                     Optional BlockInput As Boolean = False) As Integer

            ' Block Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(True)

            ' The inputs structures to send.
            Dim Inputs As New List(Of NativeMethods.Input)

            ' The current input to add into the Inputs list.
            Dim CurrentInput As New NativeMethods.Input

            ' Add a mouse movement.
            With CurrentInput
                .type = NativeMethods.InputType.Mouse
                .mi.dx = Offset.X
                .mi.dy = Offset.Y
                .mi.dwFlags = NativeMethods.MouseInput_Flags.Move
            End With : Inputs.Add(CurrentInput)

            ' Send the mouse movement.
            Return NativeMethods.SendInput(Inputs.Count, Inputs.ToArray,
                                       Marshal.SizeOf(GetType(NativeMethods.Input)))

            ' Unblock Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(False)

        End Function

        ''' <summary>
        ''' Slices the mouse position.
        ''' </summary>
        ''' <param name="X">
        ''' Indicates the 'X' offset.
        ''' </param>
        ''' <param name="Y">
        ''' Indicates the 'Y' offset.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the mouse movement is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the mouse input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function MouseMove(ByVal X As Integer, ByVal Y As Integer,
                                     Optional BlockInput As Boolean = False) As Integer

            Return MouseMove(New Point(X, Y), BlockInput)

        End Function

        ''' <summary>
        ''' Moves the mouse hotspot to an absolute position, in coordinates.
        ''' </summary>
        ''' <param name="Position">
        ''' Indicates the absolute position.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the mouse movement is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the mouse input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function MousePosition(ByVal Position As Point,
                                         Optional BlockInput As Boolean = False) As Integer

            ' Block Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(True)

            ' The inputs structures to send.
            Dim Inputs As New List(Of NativeMethods.Input)

            ' The current input to add into the Inputs list.
            Dim CurrentInput As New NativeMethods.Input

            ' Transform the coordinates.
            Position.X = CInt(Position.X * 65535 / (Screen.PrimaryScreen.Bounds.Width - 1))
            Position.Y = CInt(Position.Y * 65535 / (Screen.PrimaryScreen.Bounds.Height - 1))

            ' Add an absolute mouse movement.
            With CurrentInput
                .type = NativeMethods.InputType.Mouse
                .mi.dx = Position.X
                .mi.dy = Position.Y
                .mi.dwFlags = NativeMethods.MouseInput_Flags.Absolute Or NativeMethods.MouseInput_Flags.Move
                .mi.time = 0
            End With : Inputs.Add(CurrentInput)

            ' Send the absolute mouse movement.
            Return NativeMethods.SendInput(Inputs.Count, Inputs.ToArray,
                                       Marshal.SizeOf(GetType(NativeMethods.Input)))

            ' Unblock Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(False)

        End Function

        ''' <summary>
        ''' Moves the mouse hotspot to an absolute position, in coordinates.
        ''' </summary>
        ''' <param name="X">
        ''' Indicates the absolute 'X' coordinate.
        ''' </param>
        ''' <param name="Y">
        ''' Indicates the absolute 'Y' coordinate.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the mouse movement is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the mouse input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function MousePosition(ByVal X As Integer, ByVal Y As Integer,
                                         Optional BlockInput As Boolean = False) As Integer

            Return MousePosition(New Point(X, Y), BlockInput)

        End Function

        ''' <summary>
        ''' Simulates a mouse click.
        ''' </summary>
        ''' <param name="MouseAction">
        ''' Indicates the mouse action to perform.
        ''' </param>
        ''' <param name="BlockInput">
        ''' If set to <c>true</c>, the keyboard and mouse are blocked until the mouse movement is sent.
        ''' </param>
        ''' <returns>
        ''' The function returns the number of events that it successfully inserted into the mouse input stream. 
        ''' If the function returns zero, the input was already blocked by another thread.
        ''' </returns>
        Public Shared Function MouseClick(ByVal MouseAction As MouseButton,
                                      Optional BlockInput As Boolean = False) As Integer

            ' Block Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(True)

            ' The inputs structures to send.
            Dim Inputs As New List(Of NativeMethods.Input)

            ' The current input to add into the Inputs list.
            Dim CurrentInput As New NativeMethods.Input

            ' The mouse actions to perform.
            Dim MouseActions As New List(Of MouseButton)

            Select Case MouseAction

                Case MouseButton.LeftPress ' Left button, hold and release.
                    MouseActions.Add(MouseButton.LeftDown)
                    MouseActions.Add(MouseButton.LeftUp)

                Case MouseButton.RightPress ' Right button, hold and release.
                    MouseActions.Add(MouseButton.RightDown)
                    MouseActions.Add(MouseButton.RightUp)

                Case MouseButton.MiddlePress ' Middle button, hold and release.
                    MouseActions.Add(MouseButton.MiddleDown)
                    MouseActions.Add(MouseButton.MiddleUp)

                Case Else ' Other
                    MouseActions.Add(MouseAction)

            End Select ' MouseAction

            For Each Action As MouseButton In MouseActions

                ' Add the mouse click.
                With CurrentInput
                    .type = NativeMethods.InputType.Mouse
                    '.mi.dx = Offset.X
                    '.mi.dy = Offset.Y
                    .mi.dwFlags = Action
                End With : Inputs.Add(CurrentInput)

            Next Action

            ' Send the mouse click.
            Return NativeMethods.SendInput(Inputs.Count, Inputs.ToArray,
                                       Marshal.SizeOf(GetType(NativeMethods.Input)))

            ' Unblock Keyboard and mouse.
            If BlockInput Then NativeMethods.BlockInput(False)

        End Function

#End Region

    End Class

    Public NotInheritable Class SetWindowStyle
        ' ***********************************************************************
        ' Author           : Destroyer
        ' Last Modified On : 29-09-2020
        ' ***********************************************************************
        ' <copyright file="SetWindowStyle.vb" company="Elektro Studios">
        '     Copyright (c) All rights reserved.
        ' </copyright>
        ' ***********************************************************************

#Region " P/Invoke "

        ''' <summary>
        ''' Platform Invocation methods (P/Invoke), access unmanaged code.
        ''' This class does not suppress stack walks for unmanaged code permission.
        ''' <see cref="System.Security.SuppressUnmanagedCodeSecurityAttribute"/>  must not be applied to this class.
        ''' This class is for methods that can be used anywhere because a stack walk will be performed.
        ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/ms182161.aspx
        ''' </summary>
        Protected NotInheritable Class NativeMethods

#Region " Methods "

            ''' <summary>
            ''' Retrieves a handle to the top-level window whose class name and window name match the specified strings.
            ''' This function does not search child windows.
            ''' This function does not perform a case-sensitive search.
            ''' To search child windows, beginning with a specified child window, use the FindWindowEx function.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633499%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="lpClassName">The class name.
            ''' If this parameter is NULL, it finds any window whose title matches the lpWindowName parameter.</param>
            ''' <param name="lpWindowName">The window name (the window's title).
            ''' If this parameter is NULL, all window names match.</param>
            ''' <returns>If the function succeeds, the return value is a handle to the window that has the specified class name and window name.
            ''' If the function fails, the return value is NULL.</returns>
            <DllImport("user32.dll", SetLastError:=False, CharSet:=CharSet.Auto, BestFitMapping:=False)>
            Friend Shared Function FindWindow(
ByVal lpClassName As String,
ByVal lpWindowName As String
) As IntPtr
            End Function

            ''' <summary>
            ''' Retrieves a handle to a window whose class name and window name match the specified strings. 
            ''' The function searches child windows, beginning with the one following the specified child window. 
            ''' This function does not perform a case-sensitive search.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633500%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="hwndParent">
            ''' A handle to the parent window whose child windows are to be searched.
            ''' If hwndParent is NULL, the function uses the desktop window as the parent window. 
            ''' The function searches among windows that are child windows of the desktop. 
            ''' </param>
            ''' <param name="hwndChildAfter">
            ''' A handle to a child window. 
            ''' The search begins with the next child window in the Z order. 
            ''' The child window must be a direct child window of hwndParent, not just a descendant window.
            ''' If hwndChildAfter is NULL, the search begins with the first child window of hwndParent.
            ''' </param>
            ''' <param name="strClassName">
            ''' The window class name.
            ''' </param>
            ''' <param name="strWindowName">
            ''' The window name (the window's title). 
            ''' If this parameter is NULL, all window names match.
            ''' </param>
            ''' <returns>
            ''' If the function succeeds, the return value is a handle to the window that has the specified class and window names.
            ''' If the function fails, the return value is NULL.
            ''' </returns>
            <DllImport("User32.dll", SetLastError:=False, CharSet:=CharSet.Auto, BestFitMapping:=False)>
            Friend Shared Function FindWindowEx(
ByVal hwndParent As IntPtr,
ByVal hwndChildAfter As IntPtr,
ByVal strClassName As String,
ByVal strWindowName As String
) As IntPtr
            End Function

            ''' <summary>
            ''' Retrieves the identifier of the thread that created the specified window 
            ''' and, optionally, the identifier of the process that created the window.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633522%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="hWnd">A handle to the window.</param>
            ''' <param name="ProcessId">
            ''' A pointer to a variable that receives the process identifier. 
            ''' If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process to the variable; 
            ''' otherwise, it does not.
            ''' </param>
            ''' <returns>The identifier of the thread that created the window.</returns>
            <DllImport("user32.dll")>
            Friend Shared Function GetWindowThreadProcessId(
ByVal hWnd As IntPtr,
ByRef ProcessId As Integer
) As Integer
            End Function

            <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLong")>
            Friend Shared Function SetWindowLong32(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As Integer) As Integer
            End Function

            <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLongPtr")>
            Friend Shared Function SetWindowLongPtr64(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
            End Function

            Friend Shared Function SetWindowLongPtr(ByVal hWnd As IntPtr, nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
                If IntPtr.Size = 8 Then
                    Return SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
                Else
                    Return New IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32))
                End If
            End Function

#End Region

        End Class

#End Region

#Region " Enumerations "

        Public Enum WindowLongFlags As Integer
            GWL_EXSTYLE = -20
            GWLP_HINSTANCE = -6
            GWLP_HWNDPARENT = -8
            GWL_ID = -12
            GWL_STYLE = -16
            GWL_USERDATA = -21
            GWL_WNDPROC = -4
            DWLP_USER = &H8
            DWLP_MSGRESULT = &H0
            DWLP_DLGPROC = &H4
        End Enum

        <FlagsAttribute()>
        Public Enum WindowStyles As Long

            Todo1 = 2
            Todo2 = 2048
            Todo3 = 32768

            WS_OVERLAPPED = 0
            WS_POPUP = 2147483648
            WS_CHILD = 1073741824
            WS_MINIMIZE = 536870912
            WS_VISIBLE = 268435456
            WS_DISABLED = 134217728
            WS_CLIPSIBLINGS = 67108864
            WS_CLIPCHILDREN = 33554432
            WS_MAXIMIZE = 16777216
            WS_BORDER = 8388608
            WS_DLGFRAME = 4194304
            WS_VSCROLL = 2097152
            WS_HSCROLL = 1048576
            WS_SYSMENU = 524288
            WS_THICKFRAME = 262144
            WS_GROUP = 131072
            WS_TABSTOP = 65536

            WS_MINIMIZEBOX = 131072
            WS_MAXIMIZEBOX = 65536

            WS_CAPTION = WS_BORDER Or WS_DLGFRAME
            WS_TILED = WS_OVERLAPPED
            WS_ICONIC = WS_MINIMIZE
            WS_SIZEBOX = WS_THICKFRAME
            WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or
                      WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
            WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
            WS_CHILDWINDOW = WS_CHILD

            WS_EX_DLGMODALFRAME = 1
            WS_EX_NOPARENTNOTIFY = 4
            WS_EX_TOPMOST = 8
            WS_EX_ACCEPTFILES = 16
            WS_EX_TRANSPARENT = 32

            '#If (WINVER >= 400) Then
            WS_EX_MDICHILD = 64
            WS_EX_TOOLWINDOW = 128
            WS_EX_WINDOWEDGE = 256
            WS_EX_CLIENTEDGE = 512
            WS_EX_CONTEXTHELP = 1024

            WS_EX_RIGHT = 4096
            WS_EX_LEFT = 0
            WS_EX_RTLREADING = 8192
            WS_EX_LTRREADING = 0
            WS_EX_LEFTSCROLLBAR = 16384
            WS_EX_RIGHTSCROLLBAR = 0

            WS_EX_CONTROLPARENT = 65536
            WS_EX_STATICEDGE = 131072
            WS_EX_APPWINDOW = 262144

            WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
            WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
            '#End If

            '#If (WIN32WINNT >= 500) Then
            WS_EX_LAYERED = 524288
            '#End If

            '#If (WINVER >= 500) Then
            WS_EX_NOINHERITLAYOUT = 1048576 ' Disable inheritence of mirroring by children
            WS_EX_LAYOUTRTL = 4194304 ' Right to left mirroring
            '#End If

            '#If (WIN32WINNT >= 500) Then
            WS_EX_COMPOSITED = 33554432
            WS_EX_NOACTIVATE = 67108864
            '#End If

        End Enum

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Set the state of a window by an HWND.
        ''' </summary>
        ''' <param name="WindowHandle">A handle to the window.</param>
        ''' <param name="WindowStyle">The Style of the window.</param>
        ''' <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
        Friend Shared Function SetWindowStyle(ByVal WindowHandle As IntPtr,
                                              ByVal WindowStyle As WindowStyles) As Boolean

            Return NativeMethods.SetWindowLongPtr(WindowHandle, WindowLongFlags.GWL_STYLE, WindowStyle)

        End Function

        ''' <summary>
        ''' Set the state of a window by a process name.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process.</param>
        ''' <param name="WindowStyle">The Style of the window.</param>
        ''' <param name="Recursivity">If set to <c>false</c>, only the first process instance will be processed.</param>
        Friend Shared Sub SetWindowStyle(ByVal ProcessName As String,
                                         ByVal WindowStyle As WindowStyles,
                                         Optional ByVal Recursivity As Boolean = False)

            If ProcessName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                ProcessName = ProcessName.Remove(ProcessName.Length - ".exe".Length)
            End If

            Dim pHandle As IntPtr = IntPtr.Zero
            Dim pID As Integer = 0I

            Dim Processes As Process() = Process.GetProcessesByName(ProcessName)

            ' If any process matching the name is found then...
            If Processes.Count = 0 Then
                Exit Sub
            End If

            For Each p As Process In Processes

                ' If Window is visible then...
                If p.MainWindowHandle <> IntPtr.Zero Then
                    SetWindowStyle(p.MainWindowHandle, WindowStyle)

                Else ' Window is hidden

                    ' Check all open windows (not only the process we are looking),
                    ' begining from the child of the desktop, phandle = IntPtr.Zero initialy.
                    While pID <> p.Id ' Check all windows.

                        ' Get child handle of window who's handle is "pHandle".
                        pHandle = NativeMethods.FindWindowEx(IntPtr.Zero, pHandle, Nothing, Nothing)

                        ' Get ProcessId from "pHandle".
                        NativeMethods.GetWindowThreadProcessId(pHandle, pID)

                        ' If the ProcessId matches the "pID" then...
                        If pID = p.Id Then

                            NativeMethods.SetWindowLongPtr(pHandle, WindowLongFlags.GWL_STYLE, WindowStyle)

                            If Not Recursivity Then
                                Exit For
                            End If

                        End If

                    End While

                End If

            Next p

        End Sub

#End Region

    End Class

    Public NotInheritable Class SetWindowState

#Region " P/Invoke "

        ''' <summary>
        ''' Platform Invocation methods (P/Invoke), access unmanaged code.
        ''' This class does not suppress stack walks for unmanaged code permission.
        ''' <see cref="System.Security.SuppressUnmanagedCodeSecurityAttribute"/>  must not be applied to this class.
        ''' This class is for methods that can be used anywhere because a stack walk will be performed.
        ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/ms182161.aspx
        ''' </summary>
        Protected NotInheritable Class NativeMethods

#Region " Methods "

            ''' <summary>
            ''' Retrieves a handle to the top-level window whose class name and window name match the specified strings.
            ''' This function does not search child windows.
            ''' This function does not perform a case-sensitive search.
            ''' To search child windows, beginning with a specified child window, use the FindWindowEx function.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633499%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="lpClassName">The class name.
            ''' If this parameter is NULL, it finds any window whose title matches the lpWindowName parameter.</param>
            ''' <param name="lpWindowName">The window name (the window's title).
            ''' If this parameter is NULL, all window names match.</param>
            ''' <returns>If the function succeeds, the return value is a handle to the window that has the specified class name and window name.
            ''' If the function fails, the return value is NULL.</returns>
            <DllImport("user32.dll", SetLastError:=False, CharSet:=CharSet.Auto, BestFitMapping:=False)>
            Friend Shared Function FindWindow(
ByVal lpClassName As String,
ByVal lpWindowName As String
) As IntPtr
            End Function

            ''' <summary>
            ''' Retrieves a handle to a window whose class name and window name match the specified strings. 
            ''' The function searches child windows, beginning with the one following the specified child window. 
            ''' This function does not perform a case-sensitive search.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633500%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="hwndParent">
            ''' A handle to the parent window whose child windows are to be searched.
            ''' If hwndParent is NULL, the function uses the desktop window as the parent window. 
            ''' The function searches among windows that are child windows of the desktop. 
            ''' </param>
            ''' <param name="hwndChildAfter">
            ''' A handle to a child window. 
            ''' The search begins with the next child window in the Z order. 
            ''' The child window must be a direct child window of hwndParent, not just a descendant window.
            ''' If hwndChildAfter is NULL, the search begins with the first child window of hwndParent.
            ''' </param>
            ''' <param name="strClassName">
            ''' The window class name.
            ''' </param>
            ''' <param name="strWindowName">
            ''' The window name (the window's title). 
            ''' If this parameter is NULL, all window names match.
            ''' </param>
            ''' <returns>
            ''' If the function succeeds, the return value is a handle to the window that has the specified class and window names.
            ''' If the function fails, the return value is NULL.
            ''' </returns>
            <DllImport("User32.dll", SetLastError:=False, CharSet:=CharSet.Auto, BestFitMapping:=False)>
            Friend Shared Function FindWindowEx(
ByVal hwndParent As IntPtr,
ByVal hwndChildAfter As IntPtr,
ByVal strClassName As String,
ByVal strWindowName As String
) As IntPtr
            End Function

            ''' <summary>
            ''' Retrieves the identifier of the thread that created the specified window 
            ''' and, optionally, the identifier of the process that created the window.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633522%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="hWnd">A handle to the window.</param>
            ''' <param name="ProcessId">
            ''' A pointer to a variable that receives the process identifier. 
            ''' If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process to the variable; 
            ''' otherwise, it does not.
            ''' </param>
            ''' <returns>The identifier of the thread that created the window.</returns>
            <DllImport("user32.dll")>
            Friend Shared Function GetWindowThreadProcessId(
ByVal hWnd As IntPtr,
ByRef ProcessId As Integer
) As Integer
            End Function

            ''' <summary>
            ''' Sets the specified window's show state.
            ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx
            ''' </summary>
            ''' <param name="hwnd">A handle to the window.</param>
            ''' <param name="nCmdShow">Controls how the window is to be shown.</param>
            ''' <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
            <DllImport("User32", SetLastError:=False)>
            Friend Shared Function ShowWindow(
ByVal hwnd As IntPtr,
ByVal nCmdShow As WindowState
) As Boolean
            End Function

#End Region

        End Class

#End Region

#Region " Enumerations "

        ''' <summary>
        ''' Controls how the window is to be shown.
        ''' MSDN Documentation: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx
        ''' </summary>
        Friend Enum WindowState As Integer

            ''' <summary>
            ''' Hides the window and activates another window.
            ''' </summary>
            Hide = 0I

            ''' <summary>
            ''' Activates and displays a window. 
            ''' If the window is minimized or maximized, the system restores it to its original size and position.
            ''' An application should specify this flag when displaying the window for the first time.
            ''' </summary>
            Normal = 1I

            ''' <summary>
            ''' Activates the window and displays it as a minimized window.
            ''' </summary>
            ShowMinimized = 2I

            ''' <summary>
            ''' Maximizes the specified window.
            ''' </summary>
            Maximize = 3I

            ''' <summary>
            ''' Activates the window and displays it as a maximized window.
            ''' </summary>      
            ShowMaximized = Maximize

            ''' <summary>
            ''' Displays a window in its most recent size and position. 
            ''' This value is similar to <see cref="WindowState.Normal"/>, except the window is not actived.
            ''' </summary>
            ShowNoActivate = 4I

            ''' <summary>
            ''' Activates the window and displays it in its current size and position.
            ''' </summary>
            Show = 5I

            ''' <summary>
            ''' Minimizes the specified window and activates the next top-level window in the Z order.
            ''' </summary>
            Minimize = 6I

            ''' <summary>
            ''' Displays the window as a minimized window. 
            ''' This value is similar to <see cref="WindowState.ShowMinimized"/>, except the window is not activated.
            ''' </summary>
            ShowMinNoActive = 7I

            ''' <summary>
            ''' Displays the window in its current size and position.
            ''' This value is similar to <see cref="WindowState.Show"/>, except the window is not activated.
            ''' </summary>
            ShowNA = 8I

            ''' <summary>
            ''' Activates and displays the window. 
            ''' If the window is minimized or maximized, the system restores it to its original size and position.
            ''' An application should specify this flag when restoring a minimized window.
            ''' </summary>
            Restore = 9I

            ''' <summary>
            ''' Sets the show state based on the SW_* value specified in the STARTUPINFO structure 
            ''' passed to the CreateProcess function by the program that started the application.
            ''' </summary>
            ShowDefault = 10I

            ''' <summary>
            ''' <b>Windows 2000/XP:</b> 
            ''' Minimizes a window, even if the thread that owns the window is not responding. 
            ''' This flag should only be used when minimizing windows from a different thread.
            ''' </summary>
            ForceMinimize = 11I

        End Enum

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Set the state of a window by an HWND.
        ''' </summary>
        ''' <param name="WindowHandle">A handle to the window.</param>
        ''' <param name="WindowState">The state of the window.</param>
        ''' <returns><c>true</c> if the function succeeds, <c>false</c> otherwise.</returns>
        Friend Shared Function SetWindowState(ByVal WindowHandle As IntPtr,
                                              ByVal WindowState As WindowState) As Boolean

            Return NativeMethods.ShowWindow(WindowHandle, WindowState)

        End Function

        ''' <summary>
        ''' Set the state of a window by a process name.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process.</param>
        ''' <param name="WindowState">The state of the window.</param>
        ''' <param name="Recursivity">If set to <c>false</c>, only the first process instance will be processed.</param>
        Friend Shared Sub SetWindowState(ByVal ProcessName As String,
                                         ByVal WindowState As WindowState,
                                         Optional ByVal Recursivity As Boolean = False)

            If ProcessName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                ProcessName = ProcessName.Remove(ProcessName.Length - ".exe".Length)
            End If

            Dim pHandle As IntPtr = IntPtr.Zero
            Dim pID As Integer = 0I

            Dim Processes As Process() = Process.GetProcessesByName(ProcessName)

            ' If any process matching the name is found then...
            If Processes.Count = 0 Then
                Exit Sub
            End If

            For Each p As Process In Processes

                ' If Window is visible then...
                If p.MainWindowHandle <> IntPtr.Zero Then
                    SetWindowState(p.MainWindowHandle, WindowState)

                Else ' Window is hidden

                    ' Check all open windows (not only the process we are looking),
                    ' begining from the child of the desktop, phandle = IntPtr.Zero initialy.
                    While pID <> p.Id ' Check all windows.

                        ' Get child handle of window who's handle is "pHandle".
                        pHandle = NativeMethods.FindWindowEx(IntPtr.Zero, pHandle, Nothing, Nothing)

                        ' Get ProcessId from "pHandle".
                        NativeMethods.GetWindowThreadProcessId(pHandle, pID)

                        ' If the ProcessId matches the "pID" then...
                        If pID = p.Id Then

                            NativeMethods.ShowWindow(pHandle, WindowState)

                            If Not Recursivity Then
                                Exit For
                            End If

                        End If

                    End While

                End If

            Next p

        End Sub

#End Region

    End Class

End Namespace
