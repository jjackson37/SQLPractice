Public Class Menu

    Private Const TitleLineLength As Integer = 80

    'Text art
    Public Sub TitlePrint()
        Console.Clear()
        TitleLine(TitleLineLength)
        Console.WriteLine("IIIIII  IIIIII    I          IIIIIII  IIIIIII  IIIIII  IIIIIII")
        Console.WriteLine("I       I    I    I             I     I        I          I   ")
        Console.WriteLine("I       I    I    I             I     I        I          I   ")
        Console.WriteLine("IIIIII  I  I I    I             I     IIIIIII  IIIIII     I   ")
        Console.WriteLine("     I  I   II    I             I     I             I     I   ")
        Console.WriteLine("     I  I    II   I             I     I             I     I   ")
        Console.WriteLine("IIIIII  IIIIII I  IIIIII        I     IIIIIII  IIIIII     I   ")
        TitleLine(TitleLineLength)
    End Sub

    'Creates a line of / length is a parameter
    Private Sub TitleLine(lineLength As Integer)
        Dim i = 1
        While (i < (lineLength - 1))
            Console.Write("/")
            i += 1
        End While
        Console.WriteLine("/")
    End Sub

    'Prints the menu into the console
    Public Sub PrintMenu()
        Console.WriteLine("Welcome to my SQL test program")
        Console.WriteLine("  1. Add new record")
        Console.WriteLine("  2. Edit a current record")
        Console.WriteLine("  3. Delete a record")
        Console.WriteLine("  4. Print records")
        Console.WriteLine("Esc. Exit program")
        TitleLine(TitleLineLength)
    End Sub

End Class
