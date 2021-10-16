
Public Class Form1

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If ConsoleHost1.ManualWriter Is Nothing Then
            Console.WriteLine(" Welcome Sir... ")
        Else
            ConsoleHost1.ManualWriter.WriteLine(" Welcome Sir... ")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' //////////////////////////////////////////////////////////////// CMD Host Process
        '  Dim Proc As New System.Diagnostics.Process
        '  Proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        '  Proc.StartInfo.CreateNoWindow = False
        '  Proc.StartInfo.UseShellExecute = True
        '  Proc.StartInfo.FileName = "cmd.exe"
        '  Proc.StartInfo.Arguments = "/k @echo off & Title ConsoleHost & cls & pause >nul"
        '  Proc.Start()
        ' //////////////////////////////////////////////////////////////// CMD Host Process


        ConsoleHost1.Initialize() ' Or ConsoleHost1.Initialize(Proc) for other Program, example : CMD.EXE


        '   If you can't write to the console, then use the insecure startup method:

        '   ConsoleHost1.Unsecure_Initialize()
        ' Or
        '   ConsoleHost1.Unsecure_InitializeV2()

        If ConsoleHost1.TargetProcess IsNot Nothing Then Me.Text = "Console :" & ConsoleHost1.TargetProcess.ProcessName

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Console.ForegroundColor = ConsoleColor.White

        If ConsoleHost1.TargetProcess IsNot Nothing Then

            SendCommand(TextBox1.Text)

        Else

            If ConsoleHost1.ManualWriter Is Nothing Then
                Console.WriteLine("Command: " & TextBox1.Text)
            Else
                ConsoleHost1.ManualWriter.WriteLine("Command: " & TextBox1.Text)
            End If

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Testing()
    End Sub

    Private Sub Testing()
        For Each color As ConsoleColor In [Enum].GetValues(GetType(ConsoleColor))
            Console.ForegroundColor = color
            If ConsoleHost1.ManualWriter Is Nothing Then
                Console.WriteLine("Foreground color set to " & color)
            Else
                ConsoleHost1.ManualWriter.WriteLine("Foreground color set to " & color)
            End If
        Next

        If ConsoleHost1.ManualWriter Is Nothing Then
            Console.WriteLine("=====================================")
        Else
            ConsoleHost1.ManualWriter.WriteLine("=====================================")
        End If

        Console.ForegroundColor = ConsoleColor.White

        For Each color As ConsoleColor In [Enum].GetValues(GetType(ConsoleColor))
            Console.BackgroundColor = color

            If ConsoleHost1.ManualWriter Is Nothing Then
                Console.WriteLine("Background color set to " & color)
            Else
                ConsoleHost1.ManualWriter.WriteLine("Background color set to " & color)
            End If

        Next

        If ConsoleHost1.ManualWriter Is Nothing Then
            Console.WriteLine("=====================================")
        Else
            ConsoleHost1.ManualWriter.WriteLine("=====================================")
        End If

        Console.ResetColor()

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Try
            If ConsoleHost1.TargetProcess IsNot Nothing Then ConsoleHost1.TargetProcess.Kill()
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub

    ' For External Cosoles, like cmd :

    Public Sub SendCommand(ByVal Text As String)
        If ConsoleHost1.TargetProcess IsNot Nothing Then

            ConsoleHost1.ActiveConsoleFocus()  ' Other Method AppActivate(ConsoleHost1.TargetProcess.Id)

            Win32.Helpers.SendInputs.SendKeys(Text, True)
            Win32.Helpers.SendInputs.SendKey(Keys.Enter)

            '  Dim c As Char = Convert.ToChar(Keys.Oemtilde) ' Ñ
            '  Dim Result As Integer = SendInputs.SendKey(Convert.ToChar(c.ToString.ToLower))
            '  MessageBox.Show(String.Format("Successfull events: {0}", CStr(Result)))

            ' SendInputs.SendKey(Keys.Enter)
            ' SendInputs.SendKey(Convert.ToChar(Keys.Back))

            ' SendInputs.SendKey(Convert.ToChar(Keys.D0))
            ' SendInputs.SendKeys(Keys.Insert, BlockInput:=True)

            ' SendInputs.MouseClick(SendInputs.MouseButton.RightPress, False)
            ' SendInputs.MouseMove(5, -5)
            ' SendInputs.MousePosition(New Point(100, 500))

        End If

    End Sub

End Class
