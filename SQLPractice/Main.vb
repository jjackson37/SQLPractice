Module Main

    ''' <summary>
    ''' Main. Uses select case to run different subroutines from the records class depending on
    ''' user input.
    ''' </summary>
    Sub Main()
        Console.BufferHeight() = 2600
        Dim MainMenu As New Menu()
        Dim Record As New Records()
        Dim userInput As Char
        Do While True
            MainMenu.TitlePrint()
            MainMenu.PrintMenu()
            userInput = Console.ReadKey().KeyChar
            Console.Beep()
            Console.WriteLine()
            Select Case userInput.ToString
                Case "1"
                    Record.AddRecord()
                Case "2"
                    Record.EditRecord()
                Case "3"
                    Record.DeleteRecord()
                Case "4"
                    Record.PrintRecords()
                Case "5"
                    Record.SearchRecords()
                Case Chr(27)
                    Console.WriteLine("Goodbye!")
                    Exit Do
                Case Else
                    Console.WriteLine(userInput & " is not a valid input...")
                    Console.WriteLine("Press any key to continue...")
                    Console.ReadKey()
            End Select
        Loop
    End Sub

End Module
