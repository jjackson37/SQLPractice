Imports System.Data.SqlClient

Public Class Records

#Region "Var and Consts"
    'The DBConnectionString must be edited to the correct file directory for the program to function
    Private Const DBConnectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\jjackson\Documents\Visual Studio 2015\Projects\SQLPractice\SQLPractice\TestDB.mdf"";Integrated Security=True"
    'Shows exceptions and messages for DB connections
    Private Const Debug As Boolean = True
    Dim connection As SqlConnection
#End Region

#Region "Methods"

    'Fills connection with the DBConnectionString
    Private Sub InitializeConnection()
        connection = New SqlConnection(DBConnectionString)
    End Sub

    'Opens the DB connection
    Private Sub DBOpen()
        Try
            InitializeConnection()
            connection.Open()
            If Debug Then
                Console.WriteLine("+DB Connection opened successfully")
            End If
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(ex.Message)
            Console.ReadKey()
        End Try
    End Sub

    'Closes the DB connection
    Private Sub DBClose()
        Try
            InitializeConnection()
            connection.Close()
            connection.Dispose()
            If Debug Then
                Console.WriteLine("-DB Connection closed successfully")
            End If
            Console.ReadKey()
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(ex.Message)
            Console.ReadKey()
        End Try
    End Sub

    'Prints all records in the DB
    Public Sub PrintRecords()
        Try
            Console.WriteLine("Order by [i]d, [f]orename, [s]urname or [d]ate of birth?")
            Dim sortByCol As String = Console.ReadKey.KeyChar
            DBOpen()
            Dim command As SqlCommand
            Select Case sortByCol
                Case "i"
                    command = New SqlCommand("SELECT * FROM TestTable", connection)
                Case "f"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY ForeName", connection)
                Case "s"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY LastName", connection)
                Case "d"
                    command = New SqlCommand("SELECT * FROM TestTable ORDER BY DoB", connection)
                Case Else
                    command = New SqlCommand("SELECT * FROM TestTable", connection)
            End Select
            Dim reader As SqlDataReader = command.ExecuteReader
            Console.WriteLine("ID Forename       Surname        DoB")
            While reader.Read
                Console.Write(reader(0))
                Console.Write("  ")
                Console.Write(reader(1))
                Console.Write(reader(2))
                Console.WriteLine((reader(3)).ToShortDateString)
            End While
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    'Adds a new record into the DB
    Public Sub AddRecord()
        Dim foreName, lastName, DoB As String
        Try
            Console.WriteLine("Enter Forename:")
            foreName = Console.ReadLine()
            If ValidateUserInput(foreName) = False Then
                Throw New Exception("Invalid Forename input")
            End If
            Console.WriteLine("Enter Lastname:")
            lastName = Console.ReadLine()
            If ValidateUserInput(lastName) = False Then
                Throw New Exception("Invalid Surname input")
            End If
            Console.WriteLine("Enter DoB (dd/mm/yyyy):")
            DoB = Console.ReadLine()
            Try
                Dim DoBTest = CDate(DoB)
            Catch
                Throw New Exception("Invalid DoB input")
            End Try
            Dim format As String = "INSERT INTO TestTable( ForeName, LastName, DoB) VALUES( {0}{1}{0}, {0}{2}{0}, {0}{3}{0})"
            Dim rowInput As String = String.Format(format, Chr(39), foreName, lastName, DoB)
            DBOpen()
            Dim Command As SqlCommand = New SqlCommand(rowInput, connection)
            Command.BeginExecuteNonQuery()
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    'Searchs for a record in the DB based on a column
    Public Sub SearchRecords()
        Try
            Dim results As New ArrayList()
            Dim colName, commandString, searchTerm As String
            Console.WriteLine("Search using [f]orename or [s]urname?")
            Dim userInput As String = Console.ReadKey.KeyChar
            Console.Beep()
            Console.WriteLine()
            Select Case userInput
                Case "f"
                    colName = "ForeName"
                Case "s"
                    colName = "LastName"
                Case Else
                    Throw New Exception("Invalid selection")
            End Select
            Console.WriteLine("Enter the " & colName)
            searchTerm = Console.ReadLine()
            If ValidateUserInput(searchTerm) = False Then
                Throw New Exception("Invalid input")
            End If
            Dim format As String = "SELECT Id, ForeName, LastName, DoB FROM TestTable WHERE {1} = {0}{2}{0}"
            DBOpen()
            commandString = String.Format(format, Chr(39), colName, searchTerm)
            Dim command As SqlCommand = New SqlCommand(commandString, connection)
            Dim reader As SqlDataReader = command.ExecuteReader
            While reader.Read
                Console.Write(reader(0) & " ")
                Console.Write(reader(1) & " ")
                Console.Write(reader(2) & " ")
                Console.WriteLine((reader(3)).ToShortDateString)
                results.Add(reader(0))
            End While
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    'Deletes a record from the DB
    Public Sub DeleteRecord()
        Try
            Console.WriteLine("Enter the ID of the record you want to delete:")
            Dim idInput As Integer = Console.ReadLine()
            If ValidateUserInput(idInput.ToString) = False Then
                Throw New Exception("Invalid ID input")
            End If
            Dim deleteQuery As String = "DELETE FROM TestTable WHERE Id = " & idInput
            DBOpen()
            Dim command As SqlCommand = New SqlCommand(deleteQuery, connection)
            command.ExecuteNonQuery()
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    'Edits a record from the DB
    Public Sub EditRecord()
        Try
            Console.WriteLine("Enter the ID of the record you want to edit:")
            Dim idInput As Integer = Console.ReadLine()
            If ValidateUserInput(idInput.ToString) = False Then
                Throw New Exception("Invalid ID input")
            End If
        Catch ex As Exception
            If Debug Then
                Console.WriteLine(ex.Message)
            End If
        Finally
            DBClose()
        End Try
    End Sub

    'Validates userinput and returns true if valid and false if not
    Private Function ValidateUserInput(UserInput As String)
        Dim Valid As Boolean = True
        If UserInput.Contains(" ") Then
            Valid = False
        ElseIf UserInput = "" Then
            Valid = False
        End If
        Return Valid
    End Function

#End Region
End Class
