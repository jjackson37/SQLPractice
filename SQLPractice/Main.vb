Module Main

    Sub Main()
        Console.BufferHeight() = 800
        Dim MainMenu As New Menu()
        Dim Record As New Records()
        Dim userInput As Char
        Do While True
            MainMenu.TitlePrint()
            MainMenu.PrintMenu()
            userInput = Console.ReadKey().KeyChar
            Select Case userInput
                Case "1"
                    Record.AddRecord()
                Case "2"
                    Record.EditRecord()
                Case "3"
                    Record.DeleteRecord()
                Case "4"
                    Record.PrintRecords()
                Case Chr(27)
                    Exit Do
                Case Else
                    Console.WriteLine(" is not a valid input...")
                    Console.ReadKey()
            End Select
        Loop
    End Sub

End Module
