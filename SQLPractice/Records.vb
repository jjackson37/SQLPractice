Imports System.Data.SqlClient

Public Class Records

#Region "Var and Consts"
    Private Const DBConnectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\jjackson\Documents\Visual Studio 2015\Projects\SQLPractice\SQLPractice\TestDB.mdf"";Integrated Security=True"
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
            Console.WriteLine()
            If Debug Then
                Console.WriteLine("DB Connection opened successfully")
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
                Console.WriteLine("DB Connection closed successfully")
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
            DBOpen()
            Dim command As SqlCommand = New SqlCommand("SELECT * FROM TestTable", connection)
            Dim reader As SqlDataReader = command.ExecuteReader
            Console.WriteLine("¦Forename       ¦Surname        ¦DoB")
            While reader.Read
                Console.Write("¦")
                Console.Write(reader(1))
                Console.Write("¦")
                Console.Write(reader(2))
                Console.Write("¦")
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
        DBOpen()
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
            Console.WriteLine("Enter DoB (mm/dd/yyyy):")
            DoB = Console.ReadLine()
            Try
                Dim DoBTest = CDate(DoB)
            Catch
                Throw New Exception("Invalid DoB input")
            End Try
            Dim format As String = "INSERT INTO TestTable( ForeName, LastName, DoB) VALUES( {0}{1}{0}, {0}{2}{0}, {0}{3}{0})"
            Dim rowInput As String = String.Format(format, Chr(39), foreName, lastName, DoB)
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
    Private Function Search(input As String, column As Char)
        Dim results As ArrayList

        Return results
    End Function

    'Deletes a record from the DB
    Public Sub DeleteRecord()
        DBOpen()
        Dim searchTerm As String
        Try
            Console.WriteLine("Search using [f]orename or [s]urname?")
            Dim userInput As Char = Console.ReadKey.KeyChar
            Search(searchTerm, userInput)
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
