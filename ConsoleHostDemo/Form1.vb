
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' //////////////////////////////////////////////////////////////// CMD Host Process
        '  Dim Proc As New System.Diagnostics.Process
        '  Proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        '  Proc.StartInfo.CreateNoWindow = False
        '  Proc.StartInfo.UseShellExecute = True
        '  Proc.StartInfo.FileName = "cmd.exe"
        '  Proc.StartInfo.Arguments = "/k @echo off & Title ConsoleHost & cls " ' & pause >nul
        '  Proc.Start()
        ' //////////////////////////////////////////////////////////////// CMD Host Process


        ConsoleHostV21.Initialize() ' Or ConsoleHostV21.Initialize(Proc) for other Program, example : CMD.EXE


        '   If you can't write to the console, then use the insecure startup method:

        '  ConsoleHostV21.Unsecure_Initialize()

        If ConsoleHostV21.TargetProcess IsNot Nothing Then Me.Text = "Console :" & ConsoleHostV21.TargetProcess.ProcessName

    End Sub

#Region " GUI "

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("Command: " & TextBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Testing()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MatrixStart()
    End Sub

#End Region

#Region " Color Test "

    Private Sub Testing()
        For Each color As ConsoleColor In [Enum].GetValues(GetType(ConsoleColor))
            Console.ForegroundColor = color
            Console.WriteLine("Foreground color set to " & color)
        Next

        Console.WriteLine("=====================================")


        Console.ForegroundColor = ConsoleColor.White

        For Each color As ConsoleColor In [Enum].GetValues(GetType(ConsoleColor))
            Console.BackgroundColor = color

            Console.WriteLine("Background color set to " & color)


        Next

        Console.WriteLine("=====================================")

        Console.ResetColor()

    End Sub

#End Region

#Region " Matrix Effect "

    Dim IsStarted As Boolean = False
    Public Sub MatrixStart()
        If IsStarted = False Then
            IsStarted = True
            Dim Asynctask As New Task(New Action(Async Sub()
                                                     Console.Title = "Matrix Effect"
                                                     Console.ForegroundColor = ConsoleColor.DarkGreen
                                                     Console.WindowLeft = InlineAssignHelper(0, 0)
                                                     Console.WindowHeight = InlineAssignHelper(Console.BufferHeight, Console.LargestWindowHeight)
                                                     Console.WindowWidth = InlineAssignHelper(Console.BufferWidth, Console.LargestWindowWidth)

                                                     Console.CursorVisible = False
                                                     Dim width As Integer, height As Integer
                                                     Dim y As Integer()
                                                     Dim l As Integer()
                                                     InitializeC(width, height, y, l)
                                                     Dim ms As Integer
                                                     While True
                                                         Dim t1 As DateTime = DateTime.Now
                                                         MatrixStep(width, height, y, l)
                                                         ms = 10 - CInt(Math.Truncate(CType(DateTime.Now - t1, TimeSpan).TotalMilliseconds))
                                                         If ms > 0 Then
                                                             System.Threading.Thread.Sleep(ms)
                                                         End If
                                                         If Console.KeyAvailable Then
                                                             If Console.ReadKey().Key = ConsoleKey.F5 Then
                                                                 InitializeC(width, height, y, l)
                                                             End If
                                                         End If
                                                     End While
                                                 End Sub), TaskCreationOptions.PreferFairness)
            Asynctask.Start()
        End If

    End Sub

    Dim thistime As Boolean = False

    Private Sub MatrixStep(ByVal width As Integer, ByVal height As Integer, ByVal y As Integer(), ByVal l As Integer())
        Dim x As Integer
        thistime = Not thistime
        For x = 0 To width - 1
            If x Mod 11 = 10 Then
                If Not thistime Then
                    Continue For
                End If
                Console.ForegroundColor = ConsoleColor.White
            Else
                Console.ForegroundColor = ConsoleColor.DarkGreen
                Console.SetCursorPosition(x, inBoxY(y(x) - 2 - ((l(x) \ 40) * 2), height))
                Console.Write(R)
                Console.ForegroundColor = ConsoleColor.Green
            End If
            Console.SetCursorPosition(x, y(x))
            Console.Write(R)
            y(x) = inBoxY(y(x) + 1, height)
            Console.SetCursorPosition(x, inBoxY(y(x) - l(x), height))
            Console.Write(" "c)
        Next
    End Sub

    Private Sub InitializeC(ByRef width As Integer, ByRef height As Integer, ByRef y As Integer(), ByRef l As Integer())
        Dim h1 As Integer
        Dim h2 As Integer = (InlineAssignHelper(h1, (InlineAssignHelper(height, Console.WindowHeight)) \ 2)) \ 2
        width = Console.WindowWidth - 1
        y = New Integer(width - 1) {}
        l = New Integer(width - 1) {}
        Dim x As Integer
        Console.Clear()
        For x = 0 To width - 1
            y(x) = m_r.[Next](height)
            l(x) = m_r.[Next](h2 * (If((x Mod 11 <> 10), 2, 1)), h1 * (If((x Mod 11 <> 10), 2, 1)))
        Next
    End Sub

    Dim m_r As New Random()

    Private ReadOnly Property R() As Char
        Get
            Dim t As Integer = m_r.[Next](10)
            If t <= 2 Then
                Return ChrW(CInt(AscW("0"c)) + m_r.[Next](10))
            ElseIf t <= 4 Then
                Return ChrW(CInt(AscW("a"c)) + m_r.[Next](27))
            ElseIf t <= 6 Then
                Return ChrW(CInt(AscW("A"c) + m_r.[Next](27)))
            Else
                Return ChrW(m_r.[Next](32, 255))
            End If
        End Get
    End Property

    Public Function inBoxY(ByVal n As Integer, ByVal height As Integer) As Integer
        n = n Mod height
        If n < 0 Then
            Return n + height
        Else
            Return n
        End If
    End Function

    Private Function InlineAssignHelper(Of T)(ByRef target As T, ByVal value As T) As T
        target = value
        Return value
    End Function

#End Region

End Class
