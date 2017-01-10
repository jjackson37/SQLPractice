Public Class Menu

    Private Const TitleLineLength As Integer = 80

    ''' <summary>
    ''' Text art
    ''' </summary>
    Public Sub TitlePrint()
        Console.Clear()
        TitleLine(TitleLineLength)
        Console.WriteLine("-------QQQQQQQ--QQQQQQQ----Q----------QQQQQQQ--QQQQQQQ--QQQQQQQ--QQQQQQQ-------")
        Console.WriteLine("-------Q--------Q-----Q----Q-------------Q-----Q--------Q-----------Q----------")
        Console.WriteLine("-------Q--------Q-----Q----Q-------------Q-----Q--------Q-----------Q----------")
        Console.WriteLine("-------QQQQQQQ--Q---Q-Q----Q-------------Q-----QQQQQQQ--QQQQQQQ-----Q----------")
        Console.WriteLine("-------------Q--Q----QQ----Q-------------Q-----Q--------------Q-----Q----------")
        Console.WriteLine("-------------Q--Q-----QQ---Q-------------Q-----Q--------------Q-----Q----------")
        Console.WriteLine("-------QQQQQQQ--QQQQQQQ-Q--QQQQQQQ-------Q-----QQQQQQQ--QQQQQQQ-----Q----------")
        TitleLine(TitleLineLength)
    End Sub

    ''' <summary>
    ''' Creates a line of / length is a parameter
    ''' </summary>
    ''' <param name="lineLength">Length of line art</param>
    Private Sub TitleLine(lineLength As Integer)
        Dim i = 1
        While (i < (lineLength - 1))
            Console.Write("/")
            i += 1
        End While
        Console.WriteLine("/")
    End Sub

    ''' <summary>
    ''' Prints the menu into the console
    ''' </summary>
    Public Sub PrintMenu()
        Console.WriteLine("Welcome to my SQL test program")
        Console.WriteLine("  1. Add new record")
        Console.WriteLine("  2. Edit a current record")
        Console.WriteLine("  3. Delete a record")
        Console.WriteLine("  4. Print records")
        Console.WriteLine("  5. Search records")
        Console.WriteLine("Esc. Exit program")
        TitleLine(TitleLineLength)
    End Sub

End Class
