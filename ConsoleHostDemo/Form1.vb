Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConsoleHost1.Initialize()
    End Sub

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
        ' Console.ReadKey()
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If ConsoleHost1.TargetProcess IsNot Nothing Then ConsoleHost1.TargetProcess.Kill()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Testing()
    End Sub
End Class
